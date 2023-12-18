/**
 * libreOffice-convert: doc转pdf
 * 环境要求：预先安装libreOffice
 * 支持HTTP请求从libreOffice Container转执行换
 */
const fs = require('fs');
const path = require('path');
const url = require('url');
const async = require('async');
const tmp = require('tmp');
const superagent = require('superagent');
const { execFile } = require('child_process');

const convertWithOptions = (document, format, filter, options, callback) => {
    const tmpOptions = (options || {}).tmpOptions || {};
    const asyncOptions = (options || {}).asyncOptions || {};
    const execOptions = (options || {}).execOptions || {};
    const requestToContainer = (options || {}).requestToContainer || false;
    const containerConvertApi = (options || {}).containerConvertApi;
    const tempDir = tmp.dirSync({prefix: 'libreofficeConvert_', unsafeCleanup: true, ...tmpOptions});
    const  installDir = tmp.dirSync({prefix: 'soffice', unsafeCleanup: true, ...tmpOptions});
    const infile=path.join(tempDir.name, 'source');
    return async.auto({
        soffice: (callback) => {
            if(requestToContainer){
                return callback(null);
            }
            let paths = (options || {}).sofficeBinaryPaths || [];
            switch (process.platform) {
                case 'darwin': paths = [...paths, '/Applications/LibreOffice.app/Contents/MacOS/soffice'];
                    break;
                case 'linux': paths = [...paths, '/usr/bin/libreoffice', '/usr/bin/soffice', '/snap/bin/libreoffice'];
                    break;
                case 'win32': paths = [
                    ...paths,
                    path.join(process.env['PROGRAMFILES(X86)'], 'LIBREO~1/program/soffice.exe'),
                    path.join(process.env['PROGRAMFILES(X86)'], 'LibreOffice/program/soffice.exe'),
                    path.join(process.env.PROGRAMFILES, 'LibreOffice/program/soffice.exe'),
                ];
                    break;
                default:
                    return callback(new Error(`Operating system not yet supported: ${process.platform}`));
            }

            return async.filter(
                paths,
                (filePath, callback) => fs.access(filePath, err => callback(null, !err)),
                (err, res) => {
                    if (res.length === 0) {
                        return callback(new Error('Could not find soffice binary'));
                    }

                    return callback(null, res[0]);
                }
            );
        },
        saveSource: (callback) => fs.writeFile(infile, document, callback),
        convert: ['soffice', 'saveSource', (results, callback) => {
            let filterParam = filter?.length ? `:${filter}` : "";
            let fmt = !(filter ?? "").includes(" ") ? `${format}${filterParam}` : `"${format}${filterParam}"`;
            if(requestToContainer){
                return superagent.post(containerConvertApi)
                .send({infile, extension: fmt, outdir: tempDir.name})
                .set('Accept', 'application/json')
                .then(res=>{
                    callback(null, res.body)
                }).catch(err=>{
                    callback(err)
                })
            }else{
                let args = [];
                args.push(`-env:UserInstallation=${url.pathToFileURL(installDir.name)}`);
                args.push('--headless');
                args.push('--convert-to');
                args.push(fmt);
                args.push('--outdir');
                args.push(tempDir.name);
                args.push(path.join(tempDir.name, 'source'));
                return execFile(results.soffice, args, execOptions, callback);
            }
            
        }],
        loadDestination: ['convert', (results, callback) =>{
            if(process.env.NODE_ENV!='production' && requestToContainer)return callback(null);
            async.retry({
                times: asyncOptions.times || 3,
                interval: asyncOptions.interval || 200
            }, (callback) => fs.readFile(path.join(tempDir.name, `source.${format.split(":")[0]}`), callback), callback)
        }]
    }).then( (res) => {
        return callback(null, res.loadDestination);
    }).catch( (err) => {
        return callback(err);
    }).finally( () => {
        // remove temp file
        tempDir.removeCallback();
        installDir.removeCallback();
    });
};

const convert = (document, format, filter, opts, callback) => {
    return convertWithOptions(document, format, filter, opts||{}, callback)
};
module.exports = {
    convert,
    convertWithOptions
};
