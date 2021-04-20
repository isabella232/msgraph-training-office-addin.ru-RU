---
ms.openlocfilehash: 09ef719985094d87c438b6f98b931865dbd59c75
ms.sourcegitcommit: 8a65c826f6b229c287a782d784b6d9629aa5a3d0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/20/2021
ms.locfileid: "51899424"
---
<!-- markdownlint-disable MD002 MD041 -->

В этом упражнении вы создадим решение надстройки Office с помощью [Express.](http://expressjs.com/) Решение состоит из двух частей.

- Надстройка, реализованная в качестве статических ФАЙЛОВ HTML и JavaScript.
- Сервер Node.js/Express, обслуживающий надстройки и реализующий веб-API для получения данных для надстройки.

## <a name="create-the-server"></a>Создание сервера

1. Откройте интерфейс командной строки (CLI), перейдите в каталог, в котором необходимо создать проект, и запустите следующую команду для создания package.jsфайла.

    ```Shell
    yarn init
    ```

    Введите значения для подсказок по мере необходимости. Если вы не уверены, значения по умолчанию в порядке.

1. Запустите следующие команды для установки зависимостей.

    ```Shell
    yarn add express@4.17.1 express-promise-router@4.1.0 dotenv@8.2.0 node-fetch@2.6.1 jsonwebtoken@8.5.1@
    yarn add jwks-rsa@2.0.2 @azure/msal-node@1.0.2 @microsoft/microsoft-graph-client@2.2.1
    yarn add date-fns@2.21.1 date-fns-tz@1.1.4 isomorphic-fetch@3.0.0 windows-iana@5.0.1
    yarn add -D typescript@4.2.4 ts-node@9.1.1 nodemon@2.0.7 @types/node@14.14.41 @types/express@4.17.11
    yarn add -D @types/node-fetch@2.5.10 @types/jsonwebtoken@8.5.1 @types/microsoft-graph@1.35.0
    yarn add -D @types/office-js@1.0.174 @types/jquery@3.5.5 @types/isomorphic-fetch@0.0.35
    ```

1. Запустите следующую команду, чтобы создать tsconfig.jsфайл.

    ```Shell
    tsc --init
    ```

1. Откройте **./tsconfig.jsв** текстовом редакторе и внести следующие изменения.

    - Измените `target` значение `es6` на .
    - Отоносит `outDir` значение и установите `./dist` его.
    - Отоносит `rootDir` значение и установите `./src` его.

1. Откройте **./package.jsи** добавьте следующее свойство в JSON.

    ```json
    "scripts": {
      "start": "nodemon ./src/server.ts",
      "build": "tsc --project ./"
    },
    ```

1. Запустите следующую команду для создания и установки сертификатов разработки для надстройки.

    ```Shell
    npx office-addin-dev-certs install
    ```

    Если вам будет предложено подтверждение, подтвердим действия. После завершения команды вы увидите выход, аналогичный следующему.

    ```Shell
    You now have trusted access to https://localhost.
    Certificate: <path>\localhost.crt
    Key: <path>\localhost.key
    ```

1. Создайте новый файл **с именем .env** в корне проекта и добавьте следующий код.

    :::code language="ini" source="../demo/graph-tutorial/example.env":::

    Замените путь к localhost.crt и путь к `PATH_TO_LOCALHOST.CRT` `PATH_TO_LOCALHOST.KEY` выходу localhost.key предыдущей командой.

1. Создайте новый каталог в корне проекта с именем **src**.

1. Создайте два каталога **в каталоге ./src:** **addin** и **api.**

1. Создайте новый файл с именем **auth.ts** в **каталоге ./src/api** и добавьте следующий код.

    ```typescript
    import Router from 'express-promise-router';

    const authRouter = Router();

    // TODO: Implement this router

    export default authRouter;
    ```

1. Создайте новый файл с именем **graph.ts** в **каталоге ./src/api** и добавьте следующий код.

    ```typescript
    import Router from 'express-promise-router';

    const graphRouter = Router();

    // TODO: Implement this router

    export default graphRouter;
    ```

1. Создайте новый файл **с именем server.ts** в **каталоге ./src** и добавьте следующий код.

    :::code language="typescript" source="../demo/graph-tutorial/src/server.ts" id="ServerSnippet":::

## <a name="create-the-add-in"></a>Создание надстройки

1. Создайте новый файл **сtaskpane.html в** **каталоге ./src/addin** и добавьте следующий код.

    :::code language="html" source="../demo/graph-tutorial/src/addin/taskpane.html" id="TaskPaneHtmlSnippet":::

1. Создайте новый файл **с именем taskpane.css** в **каталоге ./src/addin** и добавьте следующий код.

    :::code language="css" source="../demo/graph-tutorial/src/addin/taskpane.css":::

1. Создайте новый файл **сtaskpane.js** в **каталоге ./src/addin** и добавьте следующий код.

    ```javascript
    // TEMPORARY CODE TO VERIFY ADD-IN LOADS
    'use strict';

    Office.onReady(info => {
      if (info.host === Office.HostType.Excel) {
        $(function() {
          $('p').text('Hello World!!');
        });
      }
    });
    ```

1. Создание нового каталога в **каталоге src/addin** с именем **активов.**

1. Добавьте три PNG-файла в этом каталоге в соответствии со следующей таблицей.

    | Имя файла   | Размер пикселей |
    |-------------|----------------|
    | icon-80.png | 80x80          |
    | icon-32.png | 32x32          |
    | icon-16.png | 16 x 16          |

    > [!NOTE]
    > Вы можете использовать любое изображение, которое необходимо для этого шага. Вы также можете скачать изображения, используемые в этом примере непосредственно [из GitHub](https://github.com/microsoftgraph/msgraph-training-office-addin/demo/graph-tutorial/src/addin/assets).

1. Создание нового каталога в корне манифеста с именем **проекта**.

1. Создайте новый файл **с именемmanifest.xml** в **папке ./manifest** и добавьте следующий код. Замените `NEW_GUID_HERE` новым GUID, например `b4fa03b8-1eb6-4e8b-a380-e0476be9e019` .

    :::code language="xml" source="../demo/graph-tutorial/manifest/manifest.xml":::

## <a name="side-load-the-add-in-in-excel"></a>Боковая загрузка надстройки в Excel

1. Запустите сервер с помощью следующей команды.

    ```Shell
    yarn start
    ```

1. Откройте браузер и просмотрите `https://localhost:3000/taskpane.html` . Вы должны увидеть `Not loaded` сообщение.

1. В браузере перейдите [в Office.com](https://www.office.com/) и войдите. Выберите **Создать** в панели инструментов левой руки, а затем выберите **таблицу**.

    ![Снимок экрана меню Create на Office.com](images/office-select-excel.png)

1. Выберите **вкладку Insert,** а затем выберите **надстройки Office.**

1. Выберите **Upload My Add-in,** а затем выберите **Просмотр**. Загрузите **файл ./manifest/manifest.xml.**

1. Выберите **кнопку "Календарь** импорта" на вкладке **Главная,** чтобы открыть задачу.

    ![Снимок экрана кнопки "Календарь импорта" на вкладке Главная](images/get-started.png)

1. После открытия области задач необходимо увидеть `Hello World!` сообщение.

    ![Снимок экрана сообщения Hello World](images/hello-world.png)
