# xlsxTemplater
**xlscTemplater** is a library that generates xlsx documents from an xlsx template by replacing pre-defined {placeholders}. It supports loop conditions, as well as the insertion of images and QRCode. The library is capable of producing output in xlsx/pdf formats.

## Requirements
- It runs on Node.js, Node>=18.
- It utilizes LibreOffice for generating PDFs. Please ensure that [LibreOffice](https://www.libreoffice.org/) is installed beforehand. Alternatively, you can use the provided Dockerfile to quickly set up a container.
  
## Usage
```

```


## Acknowledgments
The xlxsTemplater is developed based on the excellent [exceljs](https://github.com/exceljs/exceljs) library.

The pdf export functionality in xlsxTemplater depends on [LibreOffice](https://www.libreoffice.org/), drawing inspiration from the [libreoffice-convert](https://github.com/elwerene/libreoffice-convert) library.

