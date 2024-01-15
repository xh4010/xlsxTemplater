const ExcelJS = require('exceljs');
const libre = require('./lib/libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);
const QRCode = require('qrcode');
const addressRegex = /^[A-Z]+\d+$/;

const PLACEHOIDER_TYPE={
  BLANK: -1,
  VALUE: 0,
  IMAGE: 1,
  QRCODE: 2,
  INLINELOOP: 3,
}
class XlsxTemplater{
  constructor(template){
    this.template=template;
    this.vals=[];  // placeholders
    this.loops=[];  // loop placeholders
  }
  async parse(){
    this.wb = new ExcelJS.Workbook();
    this.xlsx = this.wb.xlsx;
    await this.xlsx.readFile(this.template);
    this.ws = this.wb.getWorksheet(1);
    /**
     覆盖duplicateRow方法，修复: 
     1.复制行后_merges不更新bug
     2.新行与被复制行保持相同合并的单元格
     3.被复制行首列为合并单元，仅复制合并单元格的末行，复制后仍保持首列为合并单元
    */
    this.ws.duplicateRow=duplicateRow;
    /**
     覆盖spliceRows方法，修复: 
     1.复制行后note丢失
     2.合并单元格样式错误(忽略样式)
    */
    this.ws.spliceRows=spliceRows;

    this.ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      let loopStart=false, loopKey, loopEnd=false, loopVals=[];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        let $cell=cell;
        $cell.$address=cell.address;
        if(cell.isMerged) $cell=cell.master;
        const res=parseCell($cell);
        if(!res){
          //lastColumn
          if(this.ws.getColumn(cell.col)==this.ws.lastColumn){
            if(loopStart && !loopEnd){
              throw `${$cell.address}:${loopKey}, Loop: no closure`;
            }
            if(loopKey && loopStart && loopEnd && loopVals.length>0){
              this.loops.push({
                key: loopKey,
                row: $cell.row,
                vals: loopVals
              })
            }
          }
          return;
        }
        if(loopStart){
          if(res.loopStart){
            throw `${$cell.address}, Loop: only one appearance is allowed`;
          }else if(res.loopEnd){
            if(!!res.loopKey && loopKey!=res.loopKey){
              throw `${$cell.address}:${loopKey}, Loop: key should be paired`;
            }
            loopEnd=true;
            if(loopKey && loopStart && loopEnd){
              if(res.vals.length>0){
                const vals={
                  vals: res.vals,
                  col: res.col,
                };
                if(!!res.innerLoop){
                  vals.innerLoop=res.innerLoop;
                  vals.innerLoopKey=res.innerLoopKey;
                }
                loopVals.push(vals);
              }
              if(loopVals.length>0)this.loops.push({
                key: loopKey,
                row: $cell.row,
                vals: loopVals
              })
              loopStart=false;
              loopEnd=false;
              loopKey=undefined
            }
          }else{
            if(res.vals.length>0){
              const vals={
                vals: res.vals,
                col: res.col,
              };
              loopVals.push(vals);
            }
          }
        }else{
          if(res.loopStart){
            if(loopEnd)throw `${$cell.address}, Loop: only one appearance is allowed`;
            loopStart=true;
            loopKey=res.loopKey
            if(res.vals.length>0){
              const vals={
                vals: res.vals,
                col: res.col,
              };
              loopVals.push(vals);
            }
          }else{
            //no Loop
            if(res.vals.length>0){
              const vals={
                vals: res.vals,
                row: res.row,
                col: res.col,
              };
              this.vals.push(vals);
            }
          }
        }
        //lastColumn
        if(this.ws.getColumn(cell.col)==this.ws.lastColumn){
          if(loopStart && !loopEnd){
            throw `${$cell.address}:${loopKey}, Loop: no closure`;
          }
          if(loopKey && loopStart && loopEnd && loopVals.length>0){
            this.loops.push({
              key: loopKey,
              row: $cell.row,
              vals: loopVals
            })
          }
        }

      })
    })
    // console.log('vals:',JSON.stringify(this.vals));
    // console.log('loops:',JSON.stringify(this.loops));
  }
  /**
   计算自动行高, 确保文字都显示出来
   PS: exceljs列宽与MSexcel显示列宽不一致
   */
  _autoRowHeight($cell){
    if(!$cell.text)return;

    // 1pt占单元格的宽度，以exceljs列宽近似计算得到
    const ptw=0.17481481; 
    const fontSize=$cell.font.size;

    //列宽
    let colWidth=0;
    if($cell.isMerged){
      const {model:range}=this.ws._merges[$cell.address];
      for(let i=range.left; i<=range.right; i++){
        colWidth+=this.ws.getColumn(i).width;
      }
    }else{
      colWidth+=this.ws.getColumn($cell.col).width;
    }

    //每行能容纳的中文字符数
    const charsPerLine=Math.floor(colWidth/fontSize/ptw);

    //计算行数
    const lines=$cell.text.split('\n');
    let lineCount=lines.length;
    lines.forEach(line=>{
      //中文字符数
      let chars=countChineseAndFullWidthChars(line);
      chars+=(line.length-chars)/2;
      chars=Math.ceil(chars);

      //行数
      const count=Math.ceil(chars/charsPerLine);
      if(count>0)lineCount+= count-1;
    })

    //行高近似值，内容全部显示，未考虑字间距/行距的差异
    const calcRowHeight=Math.ceil(lineCount * (fontSize + 3) + 1.4);

    const row=this.ws.getRow($cell.row);
    if(calcRowHeight > row.height)row.height=calcRowHeight;
  }
  /**
   添加图片
   */
  _addImage({filename, buffer, base64,  extension='png', cell}){
    let params;
    if(filename)params={filename, extension};
    else if(buffer)params={buffer, extension};
    else if(base64)params={base64, extension};
    if(!params)throw `addImage: filename, buffer or base64, one of them must be provided`;
    if(!cell)throw `addImage: cell must be provided`;
    let address=cell.address;
    if(cell.isMerged){
      const merge = this.ws._merges[cell.master.address];
      if(!!merge){
        const {model:range}=merge;
        address=getColumnLetter(range.left)+range.top+':'+getColumnLetter(range.right)+range.bottom;
      }
    }
    const imageId = this.wb.addImage(params);
    this.ws.addImage(imageId, address);

  }
  /**
   TODO fmt不起作用?
   */
  async render(data,{
    innerLoopSeparator='',
    separator='',
    stringify=(value) => {
      if (value instanceof Date) {
          return value.format();
      }
      return value.toString();
    },
    image=(value)=>value,
    qrcode=(value)=>value,
    qrcodeOpts={
      errorCorrectionLevel: 'H',
      margin: 1
    }
  }){
    for(let index=0; index<this.vals.length; index++){
      const cell=this.vals[index];
      const $cell=this.ws.getCell(cell.row, cell.col);

      let newVal='', fmt;
      for(let i=0; i<cell.vals.length; i++){
        const val=cell.vals[i];
        if(!data[val.key])continue;
        const dval=data[val.key]
        fmt=dval.fmt;
        if(val.type==PLACEHOIDER_TYPE.VALUE){
          let nval=stringify(dval);
          if(dval.value!==undefined)nval= stringify(dval.value);
          newVal+= (newVal==''?'':separator)+nval;
        }else if(val.type==PLACEHOIDER_TYPE.IMAGE){
          const filename=image(dval);
          this._addImage({filename, cell: $cell});
          
        }else if(val.type==PLACEHOIDER_TYPE.QRCODE){
          const code=qrcode(dval);
          const base64=await QRCode.toDataURL(code, qrcodeOpts || {});
          this._addImage({base64, cell: $cell});

        }else if(val.type==PLACEHOIDER_TYPE.INLINELOOP){
          for(let k=0; k<dval.length; k++){
            const sData=dval[k]
            for(let j=0; j<val.vals.length; j++){
              const sVal=val.vals[j];
              if(!sData[sVal.key])continue;
              if(sVal.type==PLACEHOIDER_TYPE.VALUE){
                let nval=stringify(sData[sVal.key]);
                newVal+= (k==0 && j==0?(newVal==''?'':separator):innerLoopSeparator)+ nval;
              }
            }
          }
        }
      }
      
      $cell.value=newVal;
      if(!!fmt)$cell.numFmt=fmt;
      this._autoRowHeight($cell);
    }
    for(let index=0; index<this.loops.length; index++){
      const loop=this.loops[index];
      if(!data[loop.key])return;
      let count=0;
      if(data[loop.key].length>1){
        count=data[loop.key].length-1;
        this.ws.duplicateRow(loop.row, count,true)
        for(let i=index+1; i<this.loops.length; i++){
          if(this.loops[i].row>i)this.loops[i].row+=count;
        }
      }
      for(let i=loop.row; i<=loop.row+count; i++){  //row
        const rData=data[loop.key][i-loop.row];
        for(let j=0; j<loop.vals.length; j++){  //col
          const cell=loop.vals[j];
          const col=cell.col;
          const row=i;
          let newVal='';
          let fmt;
          for(let m=0; m<cell.vals.length; m++){
            const val=cell.vals[m];
            if(!rData[val.key])continue;
            const dval=rData[val.key];
            fmt=dval.fmt;
            if(val.type==PLACEHOIDER_TYPE.VALUE){
              let nval=stringify(dval);
              if(dval.value!==undefined)nval = stringify(dval.value);
              newVal+= (newVal==''?'':separator)+nval;
            }else if(val.type==PLACEHOIDER_TYPE.INLINELOOP){
              for(let k=0; k<dval.length; k++){
                const sData=dval[k]
                for(let n=0; n<val.vals.length; n++){
                  const sVal=val.vals[n];
                  if(!sData[sVal.key])continue;
                  if(sVal.type==PLACEHOIDER_TYPE.VALUE){
                    newVal+= (k==0 && n==0?(newVal==''?'':separator):innerLoopSeparator)+stringify(sData[sVal.key]);
                  }
                }
              }
            }
          }
          const $cell=this.ws.getCell(row,col);
          $cell.value=newVal;
          if(!!fmt)$cell.numFmt=fmt;
          this._autoRowHeight($cell);
          //console.log($cell.address, $cell.style);
        }
      }
      
    }
    this.ws.name=data.name || 'xlsxTemplater';
  }
  async export(format='pdf', options=XlsxTemplater.defaultExportOption||{}){
    const buff = await this.xlsx.writeBuffer();
    if(format.toLocaleLowerCase()=='xlsx')return buff;
    const convBuff = await libre.convertAsync(buff, format, undefined, options);
    return convBuff;
  }
}
function parseCell($cell){
  const content = $cell.text.trim();
  if(!content || $cell.$parsed)return;

  const placeholderReg=/\{([^\{\}]+)\}/g;
  const imagePrefix='%image:';
  const qrcodePrefix='%qrcode:';
  
  let vals = [],
    loopKey, loopStart = false,
    loopVals=[],
    loopEnd = false,
    innerLoopStart = false,
    innerLoopEnd = false,
    innerLoopKey,
    innerLoopVals=[],
    row = $cell.row,
    col = $cell.col,
    existImage=false,
    existQrcode=false,
    flag = false;

  const matchs=content.match(placeholderReg);
  if(!matchs)return;
  for(let i=0; i<matchs.length; i++){
    let key=matchs[i].replace(/^\{/,'').replace(/\}$/,'').trim();
    if(key.startsWith('#')){
      key=key.replace(/^#/,'').trim();
      if(loopStart){
        innerLoopStart=true;
        innerLoopKey=key;
      }else if(innerLoopStart || innerLoopEnd){
        throw `${$cell.address}, innerLoop: allow nesting one`;
      }else{
        loopStart=true;
        loopKey=key;
      }
    }else if(key.startsWith('/')){
      key=key.replace(/^\//,'').trim();
      if(innerLoopStart){
        if(!!key && key!=innerLoopKey && !flag){
          throw `${$cell.address}:${key}, innerLoop: key should be paired`;
        }
      }else{
        if(!!key && loopStart && key!=loopKey){
          throw `${$cell.address}:${key}, Loop: key should be paired`;
        }
      }


      if(innerLoopStart && (!key || key==innerLoopKey)){
        innerLoopEnd=true;
        const val={
          type: PLACEHOIDER_TYPE.INLINELOOP,
          key: innerLoopKey,
          vals:innerLoopVals
        }
        vals.push(val);
      }else if(loopStart && (!key || key==loopKey)){
        loopStart=false;
        loopEnd=false;
        innerLoopStart=true;
        innerLoopEnd=true;
        innerLoopKey=loopKey;
        innerLoopVals=loopVals;
        const val={
          type: PLACEHOIDER_TYPE.INLINELOOP,
          key: innerLoopKey,
          vals:innerLoopVals
        }
        vals.push(val);
        loopKey=undefined;
        flag=true;
      }else{
        loopEnd=true;
      }
    }else if(key.startsWith(imagePrefix)){
      const reg=new RegExp(`^${imagePrefix}`);
      key=key.replace(reg, '').trim();
      const val={type: PLACEHOIDER_TYPE.IMAGE, key};
      if(!!key){
        if(innerLoopStart && !innerLoopEnd)innerLoopVals.push(val);
        else vals.push(val);
        existImage=true;
      }
    }else if(key.startsWith(qrcodePrefix)){
      const reg=new RegExp(`^${qrcodePrefix}`);
      key=key.replace(reg, '').trim();
      const val={type: PLACEHOIDER_TYPE.QRCODE, key};
      if(!!key){
        if(innerLoopStart && !innerLoopEnd)innerLoopVals.push(val);
        else vals.push(val);
        existQrcode=true;
      }
    }else{
      const val={type: PLACEHOIDER_TYPE.VALUE, key};
      if(!!key)if(innerLoopStart && !innerLoopEnd)innerLoopVals.push(val)
      else if(loopStart && !loopEnd) loopVals.push(val)
      else vals.push(val);
    }
  }
  if(innerLoopStart && !innerLoopEnd){
    throw `${$cell.address}:${innerLoopKey}, innerLoop: no closure`;
  }
  $cell.$parsed=true;
  if((existImage || existQrcode) && vals.length>1){
    throw `${$cell.address}:, image or qrcode: a cell must be exclusive`;
  }
  if((existImage || existQrcode) && innerLoopVals.length>0){
    throw `${$cell.address}:, image or qrcode: inlining is not allowed`;
  }

  if((loopStart || loopEnd)&&vals.length==0) vals.push({type: PLACEHOIDER_TYPE.BLANK});

  return {
    vals,
    row,
    col,
    loopKey,
    loopStart,
    loopEnd
  }
}

function getColumnLetter(columnNumber) {
  let columnLetter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnLetter;
}

function duplicateRow(rowNum, count, insert = false) {
  // create count duplicates of rowNum
  // either inserting new or overwriting existing rows
  const rSrc = this._rows[rowNum - 1];
  const inserts = new Array(count).fill(rSrc.values);
  this.spliceRows(rowNum + 1, insert ? 0 : count, ...inserts);

  //xh4010@163.com: 以下
  //处理合并单元格
  // 1.记录合并单元格
  const merges={};
  for (let i = 0; i < count; i++) {
    rSrc.eachCell({includeEmpty: true}, (cell, colNumber) => {
      if(cell.isMerged){
        const {model:range}=this._merges[cell.master.address];
        if(range.bottom==range.top){
          //仅处理行内合并，忽略多行合并
          if(!merges[cell.master.address])merges[cell.master.address]=[];
          merges[cell.master.address].push(colNumber);
        }
      }
    });
  }
  
  // 插入、删除操作cell.address值会更新, 而_merges[address]不会更新
  // 2.这里更新_merges, 必须倒序处理(否则可能被覆盖)
  const keys=Object.keys(this._merges).sort((a,b)=>{
    //倒序：行-列
    a=decodeAddress(a);
    b=decodeAddress(b);
    if(b.row != a.row)return b.row - a.row
    return b.col - a.col
  })
  keys.forEach(key=>{
    const {row, col}=decodeAddress(key);
    const obj=this._merges[key];
    if(obj.model.top>rowNum)obj.model.top+=count;
    if(obj.model.bottom>rowNum)obj.model.bottom+=count;
    
    if(row>rowNum){
      delete this._merges[key];
      const adress=getColumnLetter(col)+(row+count)
      this._merges[adress]=obj;
    }

  })
  // 3.第一列: 处理合并
  if(rSrc.getCell(1).isMerged){
    const rs1m=rSrc.getCell(1).master;
    const {model:range}=this._merges[rs1m.address];
    if(range.top!==range.bottom){
      const scope1=getColumnLetter(range.left)+range.top+':'+getColumnLetter(range.right)+range.bottom;
      const scope2=getColumnLetter(range.left)+range.top+':'+getColumnLetter(range.right)+(range.bottom+count);
      this.unMergeCells(scope1)
      this.mergeCells(scope2)
    }
  }
  // 4.新行合并单元格
  Object.values(merges).forEach(cols=>{
    for (let i = 0; i < count; i++) {
      //if(Math.min(...cols)==Math.max(...cols))continue;
      const scope=getColumnLetter(Math.min(...cols))+(rowNum + 1 + i)+':'+getColumnLetter(Math.max(...cols))+(rowNum + 1 + i);
      this.mergeCells(scope);
    }
  })
  //xh4010@163.com: 以上

  // now copy styles...
  for (let i = 0; i < count; i++) {
    const rDst = this._rows[rowNum + i];
    rDst.style = rSrc.style;
    rDst.height = rSrc.height;
    // eslint-disable-next-line no-loop-func
    rSrc.eachCell({includeEmpty: true}, (cell, colNumber) => {
      rDst.getCell(colNumber).style = cell.style;
    });
  }
}

function spliceRows(start, count, ...inserts) {
  // same problem as row.splice, except worse.
  const nKeep = start + count;
  const nInserts = inserts.length;
  const nExpand = nInserts - count;
  const nEnd = this._rows.length;
  let i;
  let rSrc;
  if (nExpand < 0) {
    // remove rows
    if (start === nEnd) {
      this._rows[nEnd - 1] = undefined;
    }
    for (i = nKeep; i <= nEnd; i++) {
      rSrc = this._rows[i - 1];
      if (rSrc) {
        const rDst = this.getRow(i + nExpand);
        rDst.values = rSrc.values;
        if(rSrc.note)rDst.node = rSrc.node;
        rDst.style = rSrc.style;
        rDst.height = rSrc.height;
        // eslint-disable-next-line no-loop-func
        rSrc.eachCell({includeEmpty: true}, (cell, colNumber) => {
          rDst.getCell(colNumber).style = cell.style;
        });
        this._rows[i - 1] = undefined;
      } else {
        this._rows[i + nExpand - 1] = undefined;
      }
    }
  } else if (nExpand > 0) {
    // insert new cells
    for (i = nEnd; i >= nKeep; i--) {
      rSrc = this._rows[i - 1];
      if (rSrc) {
        const rDst = this.getRow(i + nExpand);
        rDst.values = rSrc.values;
        rDst.style = rSrc.style;
        rDst.height = rSrc.height;
        // eslint-disable-next-line no-loop-func
        rSrc.eachCell({includeEmpty: true}, (cell, colNumber) => {
          rDst.getCell(colNumber).style = cell.style;
          if(!!cell.note)rDst.getCell(colNumber).note = cell.note;  //xiaohui: reserve note
          // remerge cells accounting for insert offset
          if (cell._value.constructor.name === 'MergeValue') {
            const cellToBeMerged = this.getRow(cell._row._number + nInserts).getCell(colNumber);
            const prevMaster = cell._value._master;
            const newMaster = this.getRow(prevMaster._row._number + nInserts).getCell(prevMaster._column._number);
            cellToBeMerged.merge(newMaster, true);  //xiaohui: igore style
          }
        });
      } else {
        this._rows[i + nExpand - 1] = undefined;
      }
    }
  }

  // now copy over the new values
  for (i = 0; i < nInserts; i++) {
    const rDst = this.getRow(start + i);
    rDst.style = {};
    rDst.values = inserts[i];
  }

  // account for defined names
  this.workbook.definedNames.spliceRows(this.name, start, count, nInserts);
}

// check if value looks like an address
function validateAddress(value) {
  if (!addressRegex.test(value)) {
    throw new Error(`Invalid Address: ${value}`);
  }
  return true;
}

// convert address string into structure
function decodeAddress(value) {
  let hasCol = false;
  let col = '';
  let colNumber = 0;
  let hasRow = false;
  let row = '';
  let rowNumber = 0;
  for (let i = 0, char; i < value.length; i++) {
    char = value.charCodeAt(i);
    // col should before row
    if (!hasRow && char >= 65 && char <= 90) {
      // 65 = 'A'.charCodeAt(0)
      // 90 = 'Z'.charCodeAt(0)
      hasCol = true;
      col += value[i];
      // colNumber starts from 1
      colNumber = (colNumber * 26) + char - 64;
    } else if (char >= 48 && char <= 57) {
      // 48 = '0'.charCodeAt(0)
      // 57 = '9'.charCodeAt(0)
      hasRow = true;
      row += value[i];
      // rowNumber starts from 0
      rowNumber = (rowNumber * 10) + char - 48;
    } else if (hasRow && hasCol && char !== 36) {
      // 36 = '$'.charCodeAt(0)
      break;
    }
  }
  if (!hasCol) {
    colNumber = undefined;
  } else if (colNumber > 16384) {
    throw new Error(`Out of bounds. Invalid column letter: ${col}`);
  }
  if (!hasRow) {
    rowNumber = undefined;
  }

  // in case $row$col
  value = col + row;

  const address = {
    address: value,
    col: colNumber,
    row: rowNumber,
    $col$row: `$${col}$${row}`,
  };
  return address;
}

function countChineseAndFullWidthChars(str) {
  // 匹配中文字符、全角空格和全角符号
  const chineseAndFullWidthChars = str.match(/[\u4e00-\u9fa5\u3000-\u303f\uff00-\uffef]/g);
  return chineseAndFullWidthChars ? chineseAndFullWidthChars.length : 0;
}

Date.prototype.format=function(fmt = 'yyyy-MM-dd hh:mm:ss'){
  const date=this;
  /*
  sample:
  (new Date()).format("yyyy-MM-dd hh:mm:ss.S") ==> 2006-07-02 08:09:04.423 
  (new Date()).format("yyyy-M-d h:m:s.S")      ==> 2006-7-2 8:9:4.18 
  */
  let o = {
    "M+": date.getMonth() + 1, //月份 
    "d+": date.getDate(), //日 
    "h+": date.getHours(), //小时 
    "m+": date.getMinutes(), //分 
    "s+": date.getSeconds(), //秒 
    "q+": Math.floor((date.getMonth() + 3) / 3), //季度 
    "S": date.getMilliseconds() //毫秒 
  };
  if (/(y+)/.test(fmt)) fmt = fmt.replace(RegExp.$1, (date.getFullYear() + "").substr(4 - RegExp.$1.length));
  for (let k in o)
  if (new RegExp("(" + k + ")").test(fmt)) fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
  return fmt;
}

module.exports = XlsxTemplater;