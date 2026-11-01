============================================================================
LN Library 1.0  -  LÉAME / README
============================================================================

  Español: primera parte.   English: second part.

ESPAÑOL
-------

Qué es
------
  LN ("Library New") es una biblioteca para programadores de C#: un conjunto
  de métodos listos para usar con fechas, horas, validaciones, cadenas de
  texto, envío de correo y mensajes en pantalla. No es un programa que se
  abra con doble clic: se añade como referencia a un proyecto de Visual
  Studio.
  Es gratuito y de código abierto. Este fichero explica cómo obtenerlo y por
  qué Windows puede mostrar un aviso al instalarlo, para que usted decida
  con toda la información.

Dos formas de obtenerlo: usted elige
------------------------------------
  A) El instalador ya preparado (SetupLN-es.msi, en español; hay también
     versión en inglés y multilingüe). Es lo más cómodo.
  B) El código fuente, para compilarlo usted mismo con Visual Studio
     Community, que es gratuito.
  Las dos formas dan exactamente el mismo resultado.

Sitio oficial: descárguela solo de aquí
---------------------------------------
  El único sitio oficial de LN Library es su página en GitHub, la del autor:
     https://github.com/HectorFortun2024/LN-Library
  Descargue el instalador y el código fuente solo desde esa página. La
  licencia MIT permite que otras personas copien LN Library, así que pueden
  aparecer copias en otras webs. El autor no las revisa ni las controla, y
  cualquiera podría haberlas modificado para meter algo dañino. Si encuentra
  LN Library en otro sitio, no la use: vaya al enlace oficial.

  El autor no se hace responsable de lo que pueda ocurrir con una copia
  obtenida fuera del sitio oficial, ni de no seguir las indicaciones de este
  fichero. La huella digital (explicada más abajo) le permite comprobar en
  un momento que su fichero es el original.

Por qué Windows puede mostrar un aviso
--------------------------------------
  El instalador no lleva "firma digital". Esa firma es un certificado que se
  compra cada año y que solo identifica a quien publica el programa: no
  analiza ni garantiza lo que el programa hace. Como LN Library es gratuito,
  el autor ha preferido no pagarla.

  Por eso, al descargar o abrir el instalador, el navegador o Windows
  (SmartScreen) pueden decir que el fichero "no se descarga con frecuencia"
  o que es una "aplicación no reconocida". Es el aviso habitual para
  cualquier programa sin firma, por bueno que sea: no quiere decir que se
  haya encontrado nada malo en él.

  Si quiere comprobar que el fichero es exactamente el que publicó el autor,
  sin ningún cambio, compare su huella digital (se explica más abajo). Y si
  prefiere no ejecutar el instalador, tiene la forma B: el código completo
  está a su disposición.

A) Instalar con el instalador (.msi)
------------------------------------
  Hay tres versiones del instalador, según el idioma de sus ventanas: en
  español (-es), en inglés (-en) y multilingüe (-all). Elija una; instalan
  lo mismo.
  1. Descárguelo del enlace oficial:
     https://github.com/HectorFortun2024/LN-Library/releases/latest
  2. Si el navegador avisa de que el fichero no se descarga con frecuencia,
     elija la opción de conservarlo.
  3. Haga doble clic en el fichero. Si aparece la ventana azul
     "Windows protegió su PC", pulse "Más información" y después
     "Ejecutar de todas formas".
  4. Si Windows pregunta "¿Quiere permitir que esta aplicación haga cambios
     en el dispositivo?" con "Editor: desconocido", pulse "Sí".
     "Desconocido" solo significa que no lleva firma.
  5. Siga el asistente. Dónde queda cada cosa y cómo desinstalarlo se
     explica en "LN Library - Guía de Instalación.pdf" (carpeta
     "LN (DOCUMENTOS OFICIALES)").

  Comprobar la huella digital (opcional):
     Abra PowerShell en la carpeta donde descargó el fichero y escriba (con
     el nombre del que haya descargado):
        Get-FileHash "SetupLN-es.msi"
     El valor que aparece en Hash debe ser exactamente el que corresponda:
     SetupLN-es.msi:
     B9B8B5D495BC121F926C44FDBFA0DAC559097F59FF222B94E19E87DA324EA025
     SetupLN-en.msi:
     567A2F9BF12D6FCA8829E9EE7E28A9F85F3C4D48D53B1BE8340B594FC90F474C
     SetupLN-all.msi:
     C5577944007980E9D5B7B6AA5143AE27A103F2F909D6C874641FF02A3FFC6A9B
     Si no coincide, no lo instale y descárguelo de nuevo del enlace
     oficial.

B) Compilarlo usted mismo con Visual Studio Community
-----------------------------------------------------
  Así no ejecuta nada preparado por otra persona: la biblioteca la construye
  su propio Visual Studio a partir del código, que puede leer línea a línea
  antes de compilarlo.
  1. Instale Visual Studio Community (gratuito para particulares y para
     proyectos de código abierto). Durante la instalación marque
     "Desarrollo de escritorio de .NET". Se descarga de:
     https://visualstudio.microsoft.com/vs/community/
  2. Descargue el código: en la página del repositorio, botón verde "Code" y
     luego "Download ZIP". El repositorio está en:
     https://github.com/HectorFortun2024/LN-Library
  3. Antes de descomprimirlo: clic derecho sobre el ZIP, "Propiedades",
     marque "Desbloquear" y pulse "Aceptar". Así Visual Studio no tratará
     cada fichero como venido de Internet ni pondrá pegas al compilar.
  4. Descomprímalo y abra LN\LN.sln con Visual Studio. Si pregunta si confía
     en el proyecto, es su comprobación normal para todo lo que viene de
     Internet: conteste que sí.
  5. En la barra de arriba elija "Release" y pulse "Compilar" y luego
     "Compilar solución".
  6. La biblioteca queda en LN\bin\Release\net8.0-windows8.0\LN.dll .
     Añádala como referencia a su proyecto.

  Requisito: el SDK de .NET 8 o posterior (Visual Studio 2022 versión 17.8 o
  posterior, o Visual Studio 2026).

  También puede saltarse este paso: añada el proyecto LN\LN.csproj a su
  propia solución y referéncielo como proyecto. Hay un programa de ejemplo
  en "Programa de Ejemplo de LN\Ejemplo\Ejemplo.sln".

Licencia
--------
  Licencia MIT: puede usarla, copiarla, modificarla y distribuirla
  libremente, también en programas comerciales. No usa paquetes de terceros.
  Texto completo en "LN (DOCUMENTOS OFICIALES)\End User Agreement.rtf".


============================================================================

ENGLISH
-------

What it is
----------
  LN ("Library New") is a library for C# programmers: a set of ready-to-use
  methods for dates, times, validations, text strings, sending e-mail and
  on-screen messages. It is not a program you open with a double click: you
  add it as a reference to a Visual Studio project.
  It is free and open source. This file explains how to get it and why
  Windows may show a warning when you install it, so you can decide with all
  the facts.

Two ways to get it: your choice
-------------------------------
  A) The ready-made installer (SetupLN-en.msi, in English; there are also
     Spanish and multilingual versions). The easiest way.
  B) The source code, to build it yourself with Visual Studio Community,
     which is free.
  Both ways give exactly the same result.

Official site: download it only from here
-----------------------------------------
  The only official site for LN Library is its GitHub page, the author's
  own:
     https://github.com/HectorFortun2024/LN-Library
  Download the installer and the source code only from that page. The MIT
  license allows other people to copy LN Library, so copies may appear on
  other websites. The author does not review or control them, and anyone
  could have modified them to add something harmful. If you find LN Library
  anywhere else, do not use it: go to the official link.

  The author is not responsible for anything that may happen with a copy
  obtained outside the official site, or from not following the instructions
  in this file. The digital fingerprint (explained below) lets you check in
  a moment that your file is the original.

Why Windows may show a warning
------------------------------
  The installer has no "digital signature". That signature is a certificate
  bought every year that only identifies who publishes the program: it does
  not analyze or guarantee what the program does. Since LN Library is free,
  the author chose not to pay for it.

  That is why, when you download or open the installer, the browser or
  Windows (SmartScreen) may say the file "isn't commonly downloaded" or is
  an "unrecognized app". This is the usual warning for any unsigned program,
  however good it is: it does not mean anything bad was found in it.

  If you want to check that the file is exactly the one the author
  published, unchanged, compare its digital fingerprint (explained below).
  And if you prefer not to run the installer, you have way B: the complete
  source code is available to you.

A) Install with the installer (.msi)
------------------------------------
  There are three versions of the installer, depending on the language of
  its windows: Spanish (-es), English (-en) and multilingual (-all). Choose
  one; they install the same thing.
  1. Download it from the official link:
     https://github.com/HectorFortun2024/LN-Library/releases/latest
  2. If the browser warns that the file isn't commonly downloaded, choose
     the option to keep it.
  3. Double-click the file. If the blue window "Windows protected your PC"
     appears, click "More info" and then "Run anyway".
  4. If Windows asks "Do you want to allow this app to make changes to your
     device?" with "Publisher: Unknown", click "Yes". "Unknown" only means
     it is not signed.
  5. Follow the wizard. Where everything goes and how to uninstall it is
     explained in "LN Library - Installation Guide.pdf" (folder
     "LN (DOCUMENTOS OFICIALES)").

  Checking the digital fingerprint (optional):
     Open PowerShell in the folder where you downloaded the file and type
     (with the name of the one you downloaded):
        Get-FileHash "SetupLN-en.msi"
     The value shown under Hash must be exactly the matching one:
     SetupLN-es.msi:
     B9B8B5D495BC121F926C44FDBFA0DAC559097F59FF222B94E19E87DA324EA025
     SetupLN-en.msi:
     567A2F9BF12D6FCA8829E9EE7E28A9F85F3C4D48D53B1BE8340B594FC90F474C
     SetupLN-all.msi:
     C5577944007980E9D5B7B6AA5143AE27A103F2F909D6C874641FF02A3FFC6A9B
     If it does not match, do not install it; download it again from the
     official link.

B) Build it yourself with Visual Studio Community
-------------------------------------------------
  This way you run nothing prepared by someone else: your own Visual Studio
  builds the library from the source code, which you can read line by line
  before building it.
  1. Install Visual Studio Community (free for individuals and for open
     source projects). During setup, tick ".NET desktop development".
     Download it from:
     https://visualstudio.microsoft.com/vs/community/
  2. Download the code: on the repository page, green "Code" button, then
     "Download ZIP". The repository is at:
     https://github.com/HectorFortun2024/LN-Library
  3. Before unzipping it: right-click the ZIP, "Properties", tick "Unblock"
     and click "OK". This way Visual Studio will not treat every file as
     coming from the Internet or object when building.
  4. Unzip it and open LN\LN.sln with Visual Studio. If it asks whether you
     trust the project, that is its normal check for anything coming from
     the Internet: answer yes.
  5. In the top bar choose "Release" and click "Build", then
     "Build Solution".
  6. The library is in LN\bin\Release\net8.0-windows8.0\LN.dll . Add it as a
     reference to your project.

  Requirement: the .NET 8 SDK or later (Visual Studio 2022 version 17.8 or
  later, or Visual Studio 2026).

  You can also skip this step: add the LN\LN.csproj project to your own
  solution and reference it as a project. There is a sample program in
  "Programa de Ejemplo de LN\Ejemplo\Ejemplo.sln".

License
-------
  MIT license: you may use, copy, modify and distribute it freely, including
  in commercial programs. It uses no third-party packages. Full text in "LN
  (DOCUMENTOS OFICIALES)\End User Agreement.rtf".


----------------------------------------------------------------------------
Autor / Author: Héctor Fortún Castañón
