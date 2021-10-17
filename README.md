# ChrTranEncryptDecryptClass

> Clase VBA para encriptación de textos.

## Definición
**ChrTranEncryptDecryptClass** es un Módulo de Clase VBA (.cls) que permite encriptar/desencriptar cadenas de texto.

> Es necesario tener conocimientos básicos de programación de Clases VBA (POO, Programación Orientada a Objetos).

Trabaja bajo MS Excel(c) versión 2007 o superior. No requiere instalación, sólo debe ser importardo a su Proyecto VBA, crear una instancia de objeto y luego usar su interfaz de métodos.

## Modo de uso
  1.  Descargue el archivo [ChrTranEncryptDecryptClass.cls](./project-dist/ChrTranEncryptDecryptClass_v1.0.0.cls).
  2.  Cree un nuevo libro habilitado para macros de Excel.
  3.  Abra el Editor de Proyectos VBA con **Ctrl+F11**
  4.  Vaya al menú ***Archivo > Importar Archivo*** o presione ***Ctrl+M*** y luego busque su archivo **ChrTranEncryptDecryptClass.cls** descargado e impórtelo al Proyecto VBA.
  5.  Cree una instancia de ojeto y llame a uno de los dos métodos de ChrTranEncryptDecryptClass y...
      ```vb
        Dim MyEncDec as ChrTranEncryptDecryptClass
        Dim MyText as String, MyEncryptedText as String

        let MyText = "Some text to encrypt"
        set MyEncDec = New ChrTranEncryptDecryptClass

        let MyEncryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyText) ' MyEncryptedText contiene ahora el texto encriptado.
        ' ...
      ```
  6. ¡Disfrute de **ChrTranEncryptDecryptClass**!

## A cerca de los dos métodos de ChrTranEncryptDecryptClass
  1. ***El Método CHRTRANGenerateRandomKey:*** Permite obtener una cadena de hasta diez caracteres aleatorios que pueden servir como claves de encriptación. Posee tres parámetros opcionales:

      1.  ***LowWord***: un número (mayor o igual a 65 y menor a 90) que corresponde al caracter inicial del rango de caracteres ASCII.
      2.  ***UpWord***: un número (menor o igual a 90 y mayor a 65) que corresponde al caracter final del rango de caracteres ASCII.
      3.  ***xLenght***: una constante: *ctedKeyLengthRandom*, para un número aleatorio de caracteres entre 1 y 10 ó *ctedKeyLength1 ... ctedKeyLength10*. Por defecto: ctedKeyLength5 (cadena de cinco caracteres).

      Ejemplo de uso:
      ```vb
        Dim MyEncDec as ChrTranEncryptDecryptClass
        Dim MyRndText as String

        set MyEncDec = New ChrTranEncryptDecryptClass
        let MyRndText = MyEncDec.CHRTRANGenerateRandomKey() ' MyRndText contiene ahora una cadena con caracteres aleatorios.
        ' ...
      ```
  2. ***El Método CHRTRANEncryptDecrypt:*** Permite encriptar o desencriptar una cadena de texto. Posee seis parámetros, de los cuales cinco son opcionales:

      1.  ***UserText***: el texto a encriptar/desencriptar.
      2.  ***UserKey***: Opcional. El texto llave que permite encriptar/desencriptar *UserText*. Si omite este valor, el algoritmo de la clase generará una llave aleatoria con una cantidad entre 1 y 10 caracteres aleatorios. Si la llave de encriptación fue generada por el algoritmo, *UserKey* no se requiere para desencriptar.
      3.  ***EncDec***: Opcional. Una constante: *ctedEncrypt*, para un encriptar, o *ctedDecrypt*, para desencriptar. Por defecto: *ctedDecrypt*.
      4.  ***RndKeyLength***: Opcional. una constante que determina la cantidad de caracteres que tendrá la llave aleatoria que debe generar el algoritmo, posibles valores: *ctedKeyLengthRandom*, para un número aleatorio de caracteres entre 1 y 10 ó *ctedKeyLength1 ... ctedKeyLength10*. Por defecto: ctedKeyLength5 (cadena de cinco caracteres). No se requiere para desencriptar.
      5.  ***RndKeyPosition***: Opcional. una constante que determina la posición que tendrá la llave aleatoria generada por el algoritmo en la cadena resultante de la encriptación. No se requiere para desencriptar. Posibles valores: *ctedRandom* (por defecto), para una posición aleatoria, ó *ctedLeft, ctedRight*: izquierda o derecha de la cadena resultante.
      5.  ***ResultType***: Opcional. una constante que determina el tipo de resultado de la cadena encriptada. Posibles valores: *ctedStatic* (por defecto), obtendrá una cadena encriptada con valores fijos, ó *ctedDymamic*, la cadena resultante contendrá caracteres aleatorios siempre; si fue utilizada en la encriptación, será requerida para la desencriptación.
      
      Ejemplo de uso:
      ```vb
        Dim MyEncDec as ChrTranEncryptDecryptClass
        Dim MyText as String, MyKey as String, MyEncryptedText as String, MyDecryptedText as String

        set MyEncDec = New ChrTranEncryptDecryptClass
        let MyKey = "myk3y"
        let MyText = "Some text to encrypt"
        
        ' Encriptado/Desencriptado con una llave propia. Resultado estático.
        let MyEncryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyText, MyKey) ' GOHHIHHGLHGFGJGHHKHGFHHOHHKGJGHHKHHFGJGHGFHGOHFNHHIHIFHHGHHK
        let MyDecryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyEncryptedText, MyKey) ' "Some text to encrypt"
        
        ' Encriptado/Desencriptado con una llave aleatoria. Resultado estático.
        let MyEncryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyText) ' PFOHJ#$GLIGNGGNNGNGGGHGOLGNGHFFGOLGGHGOLGOGGGHGNGGOFGMOGOJHFGGOHGOL
        let MyDecryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyEncryptedText) ' "Some text to encrypt"

        ' Encriptado/Desencriptado con una llave propia. Resultado dinámico.
        let MyEncryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyText, MyKey, ResultType:=ctedDymamic) ' owppqppotponoroppsponppwppsoroppsppnoroponpowpnvppqpqnppopps
        let MyDecryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyEncryptedText, MyKey, ResultType:=ctedDymamic) ' "Some text to encrypt"
        
        ' Encriptado/Desencriptado con una llave aleatoria. Resultado dinámico.
        let MyEncryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyText, ResultType:=ctedDymamic) ' #TRLWR$QVWQYSQXUQXUQQVRPPQXURPTRPPQQVRPPQYUQQVQXUQYTQXSQYXRPUQYVRPP
        let MyDecryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyEncryptedText, ResultType:=ctedDymamic) ' "Some text to encrypt"

        ' Encriptado/Desencriptado con una llave aleatoria. Resultado estático, posición de llave izquierda.
        let MyEncryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyText, RndKeyPosition:=ctedLeft) ' #TRLWR$QVWQYSQXUQXUQQVRPPQXURPTRPPQQVRPPQYUQQVQXUQYTQXSQYXRPUQYVRPP
        let MyDecryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyEncryptedText) ' "Some text to encrypt"
        
        ' Encriptado/Desencriptado con una llave aleatoria. Resultado estático, posición de llave derecha.
        let MyEncryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyText, RndKeyPosition:=ctedRight) ' QUXQYWQXQQWVQPWQYQQWVQYUQYQQPWQYQQXVQPWQWVQXUQWTQXYQYVQXWQYQ$KVHKI#
        let MyDecryptedText = MyEncDec.CHRTRANEncryptDecrypt(MyEncryptedText) ' "Some text to encrypt"    

        set MyEncDec = nothing
      ```

## Colaborar en GitHub:
El código fuente de **ChrTranEncryptDecryptClass** está en: [el directorio project-dev](./project-dev/ChrTranEncryptDecryptClass.cls) del repositorio oficial.

Tan pronto como se descargue, puede colaborar con mejoras en el algoritmo siempre bajo el respeto de [Términos de licencia](./LICENSE), [El Código de Conducta](./CODE_OF_CONDUCT.md) y los [Términos de Contribución](./CONTRIBUTING.md).

## Sitio Web

[ChrTranEncryptDecryptClass](https://roccouu.github.io/ChrTranEncryptDecryptClass/docs/index.html)

## Tutorial

[Tutorial ChrTranEncryptDecryptClass](https://roccouu.github.io/ChrTranEncryptDecryptClass/docs/index.html#/tutorial)

## Documentación

[Documentación ChrTranEncryptDecryptClass](https://roccouu.github.io/ChrTranEncryptDecryptClass/index.html#/docs/index.html#/documentation)

## Contribución

Vea las [Guías de CONTRIBUCIÓN](./CONTRIBUTING.md)

## English Readme

[README-EN.md](./README-EN.md)

## Licencia

[MIT](./LICENSE) © | [Roccou](rocky.romay@gmail.com) | 2020 - 2021 | Potosí - Bolívia