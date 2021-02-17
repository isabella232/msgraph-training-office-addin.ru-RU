---
ms.openlocfilehash: 70df2dfceb1d2df527acfecab894b28019c6febb
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274382"
---
<!-- markdownlint-disable MD002 MD041 -->

В этом упражнении вы создадим решение для надстройки Office с помощью [Express.](http://expressjs.com/) Решение состоит из двух частей.

- Надстройка, реализованная в качестве статических файлов HTML и JavaScript.
- Сервер Node.js/Express, обслуживающий надстройки и реализующий веб-API для извлечения данных для надстройки.

## <a name="create-the-server"></a>Создание сервера

1. Откройте интерфейс командной строки (CLI), перейдите в каталог, в котором нужно создать проект, и запустите следующую команду, чтобы создать package.jsв файле.

    ```Shell
    yarn init
    ```

    Введите значения для соответствующих подсказок. Если вы не уверены, значения по умолчанию вполне полнее.

1. Чтобы установить зависимости, запустите следующие команды.

    ```Shell
    yarn add express@4.17.1 express-promise-router@4.0.1 dotenv@8.2.0 node-fetch@2.6.1 jsonwebtoken@8.5.1@
    yarn add jwks-rsa@1.11.0 @azure/msal-node@1.0.0-beta.1 @microsoft/microsoft-graph-client@2.1.1
    yarn add date-fns@2.16.1 date-fns-tz@1.0.12 isomorphic-fetch@3.0.0 windows-iana@4.2.1
    yarn add -D typescript@4.0.5 ts-node@9.0.0 nodemon@2.0.6 @types/node@14.14.7 @types/express@4.17.9
    yarn add -D @types/node-fetch@2.5.7 @types/jsonwebtoken@8.5.0 @types/microsoft-graph@1.26.0
    yarn add -D @types/office-js@1.0.147 @types/jquery@3.5.4 @types/isomorphic-fetch@0.0.35
    ```

1. Чтобы создать файл с tsconfig.js, tsconfig.jsследующую команду.

    ```Shell
    tsc --init
    ```

1. Откройте **./tsconfig.jsв** текстовом редакторе и внести следующие изменения.

    - Измените `target` значение на `es6` .
    - Разкомментив `outDir` значение и запланив для него значение `./dist` .
    - Разкомментив `rootDir` значение и запланив для него значение `./src` .

1. Откройте **./package.jsи** добавьте следующее свойство в JSON.

    ```json
    "scripts": {
      "start": "nodemon ./src/server.ts",
      "build": "tsc --project ./"
    },
    ```

1. Чтобы создать и установить сертификаты разработки для надстройки, запустите следующую команду.

    ```Shell
    npx office-addin-dev-certs install
    ```

    Если вам будет предложено подтвердить, подтвердит действия. После выполнения команды вы увидите выходные данные, похожие на следующие.

    ```Shell
    You now have trusted access to https://localhost.
    Certificate: <path>\localhost.crt
    Key: <path>\localhost.key
    ```

1. Создайте файл **с именем ENV** в корневой папке проекта и добавьте следующий код.

    :::code language="ini" source="../demo/graph-tutorial/example.env":::

    Замените путь к localhost.crt и путь к `PATH_TO_LOCALHOST.CRT` `PATH_TO_LOCALHOST.KEY` выходу localhost.key предыдущей командой.

1. Создайте каталог в корневой каталог проекта с именем **src.**

1. Создайте два каталога в **каталоге ./src:** **надстройку** и **API.**

1. Создайте файл с именем **auth.ts** в **каталоге ./src/api** и добавьте следующий код.

    ```typescript
    import Router from 'express-promise-router';

    const authRouter = Router();

    // TODO: Implement this router

    export default authRouter;
    ```

1. Создайте файл **graph.ts** в **каталоге ./src/api** и добавьте следующий код.

    ```typescript
    import Router from 'express-promise-router';

    const graphRouter = Router();

    // TODO: Implement this router

    export default graphRouter;
    ```

1. Создайте файл с **именем server.ts** в **каталоге ./src** и добавьте следующий код.

    :::code language="typescript" source="../demo/graph-tutorial/src/server.ts" id="ServerSnippet":::

## <a name="create-the-add-in"></a>Создание надстройки

1. Создайте файл с **именемtaskpane.html** в **каталоге ./src/addin** и добавьте следующий код.

    :::code language="html" source="../demo/graph-tutorial/src/addin/taskpane.html" id="TaskPaneHtmlSnippet":::

1. Создайте файл **taskpane.css** в **каталоге ./src/addin** и добавьте следующий код.

    :::code language="css" source="../demo/graph-tutorial/src/addin/taskpane.css":::

1. Создайте файл с **именемtaskpane.js** в **каталоге ./src/addin** и добавьте следующий код.

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

1. Создайте новый каталог в каталоге **.src/addin** с именем **assets.**

1. Добавьте три файла PNG в этот каталог согласно следующей таблице.

    | Имя файла   | Размер в пикселях |
    |-------------|----------------|
    | icon-80.png | 80x80          |
    | icon-32.png | 32x32          |
    | icon-16.png | 16 x 16          |

    > [!NOTE]
    > Для этого шага можно использовать любое изображение. Вы также можете скачать изображения, используемые в этом примере, непосредственно с [GitHub.](https://github.com/microsoftgraph/msgraph-training-office-addin/demo/graph-tutorial/src/addin/assets)

1. Создайте каталог в корневом каталоге **манифеста** проекта с именем .

1. Создайте файл с **именемmanifest.xml** в **папке ./manifest** и добавьте следующий код. Замените `NEW_GUID_HERE` новым GUID, например `b4fa03b8-1eb6-4e8b-a380-e0476be9e019` .

    :::code language="xml" source="../demo/graph-tutorial/manifest/manifest.xml":::

## <a name="side-load-the-add-in-in-excel"></a>Загрузка неогрузки надстройки в Excel

1. Запустите сервер с помощью следующей команды.

    ```Shell
    yarn start
    ```

1. Откройте браузер и перейдите к `https://localhost:3000/taskpane.html` . Должно отобраться `Not loaded` сообщение.

1. В браузере перейдите в [Office.com](https://www.office.com/) и войдите. Выберите **"Создать"** на левой панели инструментов, а затем выберите **"Таблица".**

    ![Снимок экрана: меню "Создать" Office.com](images/office-select-excel.png)

1. Выберите **вкладку "Вставка"** и выберите **надстройки Office.**

1. Выберите **"Отправить мою надстройка"** и выберите **"Обзор".** Загрузите **файл ./manifest/manifest.xml.**

1. Выберите **кнопку "Импорт календаря"** на **вкладке "Главная",** чтобы открыть ее.

    ![Снимок экрана с кнопкой "Импорт календаря" на вкладке "Главная"](images/get-started.png)

1. После открытия области задач должно отобраться `Hello World!` сообщение.

    ![Снимок экрана: сообщение Hello World](images/hello-world.png)
