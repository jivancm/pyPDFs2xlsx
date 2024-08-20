# pyPDF2xlsx
Script básico y simple para extraer información de archivos PDF, en particular de los acuses de declaración patrimonial y exportar contenido de este documento a un archivo de Excel.

## Uso
> Carpeta donde se buscarán los archivos PDF

      gsDir = './EMAIL'

> Carpeta donde se moverán los archivos PDF en los que se encuentren las cadenas buscadas

    gsOkDir = gsDir + '/OK'

> Carpeta donde se moverán los archivos PDF en los que no se encuentren las cadenas buscadas

    gsErrDir = gsDir + '/ERROR'

> Nombre del libro de Excel en el que se guardarán los datos

    excel = "./FORMATO.xlsx"
