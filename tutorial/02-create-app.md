---
ms.openlocfilehash: 70df2dfceb1d2df527acfecab894b28019c6febb
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274382"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="43b43-101">В этом упражнении вы создадим решение для надстройки Office с помощью [Express.](http://expressjs.com/)</span><span class="sxs-lookup"><span data-stu-id="43b43-101">In this exercise you will create an Office Add-in solution using [Express](http://expressjs.com/).</span></span> <span data-ttu-id="43b43-102">Решение состоит из двух частей.</span><span class="sxs-lookup"><span data-stu-id="43b43-102">The solution will consist of two parts.</span></span>

- <span data-ttu-id="43b43-103">Надстройка, реализованная в качестве статических файлов HTML и JavaScript.</span><span class="sxs-lookup"><span data-stu-id="43b43-103">The add-in, implemented as static HTML and JavaScript files.</span></span>
- <span data-ttu-id="43b43-104">Сервер Node.js/Express, обслуживающий надстройки и реализующий веб-API для извлечения данных для надстройки.</span><span class="sxs-lookup"><span data-stu-id="43b43-104">A Node.js/Express server that serves the add-in and implements a web API to retrieve data for the add-in.</span></span>

## <a name="create-the-server"></a><span data-ttu-id="43b43-105">Создание сервера</span><span class="sxs-lookup"><span data-stu-id="43b43-105">Create the server</span></span>

1. <span data-ttu-id="43b43-106">Откройте интерфейс командной строки (CLI), перейдите в каталог, в котором нужно создать проект, и запустите следующую команду, чтобы создать package.jsв файле.</span><span class="sxs-lookup"><span data-stu-id="43b43-106">Open your command-line interface (CLI), navigate to a directory where you want to create your project, and run the following command to generate a package.json file.</span></span>

    ```Shell
    yarn init
    ```

    <span data-ttu-id="43b43-107">Введите значения для соответствующих подсказок.</span><span class="sxs-lookup"><span data-stu-id="43b43-107">Enter values for the prompts as appropriate.</span></span> <span data-ttu-id="43b43-108">Если вы не уверены, значения по умолчанию вполне полнее.</span><span class="sxs-lookup"><span data-stu-id="43b43-108">If you're unsure, the default values are fine.</span></span>

1. <span data-ttu-id="43b43-109">Чтобы установить зависимости, запустите следующие команды.</span><span class="sxs-lookup"><span data-stu-id="43b43-109">Run the following commands to install dependencies.</span></span>

    ```Shell
    yarn add express@4.17.1 express-promise-router@4.0.1 dotenv@8.2.0 node-fetch@2.6.1 jsonwebtoken@8.5.1@
    yarn add jwks-rsa@1.11.0 @azure/msal-node@1.0.0-beta.1 @microsoft/microsoft-graph-client@2.1.1
    yarn add date-fns@2.16.1 date-fns-tz@1.0.12 isomorphic-fetch@3.0.0 windows-iana@4.2.1
    yarn add -D typescript@4.0.5 ts-node@9.0.0 nodemon@2.0.6 @types/node@14.14.7 @types/express@4.17.9
    yarn add -D @types/node-fetch@2.5.7 @types/jsonwebtoken@8.5.0 @types/microsoft-graph@1.26.0
    yarn add -D @types/office-js@1.0.147 @types/jquery@3.5.4 @types/isomorphic-fetch@0.0.35
    ```

1. <span data-ttu-id="43b43-110">Чтобы создать файл с tsconfig.js, tsconfig.jsследующую команду.</span><span class="sxs-lookup"><span data-stu-id="43b43-110">Run the following command to generate a tsconfig.json file.</span></span>

    ```Shell
    tsc --init
    ```

1. <span data-ttu-id="43b43-111">Откройте **./tsconfig.jsв** текстовом редакторе и внести следующие изменения.</span><span class="sxs-lookup"><span data-stu-id="43b43-111">Open **./tsconfig.json** in a text editor and make the following changes.</span></span>

    - <span data-ttu-id="43b43-112">Измените `target` значение на `es6` .</span><span class="sxs-lookup"><span data-stu-id="43b43-112">Change the `target` value to `es6`.</span></span>
    - <span data-ttu-id="43b43-113">Разкомментив `outDir` значение и запланив для него значение `./dist` .</span><span class="sxs-lookup"><span data-stu-id="43b43-113">Uncomment the `outDir` value and set it to `./dist`.</span></span>
    - <span data-ttu-id="43b43-114">Разкомментив `rootDir` значение и запланив для него значение `./src` .</span><span class="sxs-lookup"><span data-stu-id="43b43-114">Uncomment the `rootDir` value and set it to `./src`.</span></span>

1. <span data-ttu-id="43b43-115">Откройте **./package.jsи** добавьте следующее свойство в JSON.</span><span class="sxs-lookup"><span data-stu-id="43b43-115">Open **./package.json** and add the following property to the JSON.</span></span>

    ```json
    "scripts": {
      "start": "nodemon ./src/server.ts",
      "build": "tsc --project ./"
    },
    ```

1. <span data-ttu-id="43b43-116">Чтобы создать и установить сертификаты разработки для надстройки, запустите следующую команду.</span><span class="sxs-lookup"><span data-stu-id="43b43-116">Run the following command to generate and install development certificates for your add-in.</span></span>

    ```Shell
    npx office-addin-dev-certs install
    ```

    <span data-ttu-id="43b43-117">Если вам будет предложено подтвердить, подтвердит действия.</span><span class="sxs-lookup"><span data-stu-id="43b43-117">If prompted for confirmation, confirm the actions.</span></span> <span data-ttu-id="43b43-118">После выполнения команды вы увидите выходные данные, похожие на следующие.</span><span class="sxs-lookup"><span data-stu-id="43b43-118">Once the command completes, you will see output similar to the following.</span></span>

    ```Shell
    You now have trusted access to https://localhost.
    Certificate: <path>\localhost.crt
    Key: <path>\localhost.key
    ```

1. <span data-ttu-id="43b43-119">Создайте файл **с именем ENV** в корневой папке проекта и добавьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="43b43-119">Create a new file named **.env** in the root of your project and add the following code.</span></span>

    :::code language="ini" source="../demo/graph-tutorial/example.env":::

    <span data-ttu-id="43b43-120">Замените путь к localhost.crt и путь к `PATH_TO_LOCALHOST.CRT` `PATH_TO_LOCALHOST.KEY` выходу localhost.key предыдущей командой.</span><span class="sxs-lookup"><span data-stu-id="43b43-120">Replace `PATH_TO_LOCALHOST.CRT` with the path to localhost.crt and `PATH_TO_LOCALHOST.KEY` with the path to localhost.key output by the previous command.</span></span>

1. <span data-ttu-id="43b43-121">Создайте каталог в корневой каталог проекта с именем **src.**</span><span class="sxs-lookup"><span data-stu-id="43b43-121">Create a new directory in the root of your project named **src**.</span></span>

1. <span data-ttu-id="43b43-122">Создайте два каталога в **каталоге ./src:** **надстройку** и **API.**</span><span class="sxs-lookup"><span data-stu-id="43b43-122">Create two directories in the **./src** directory: **addin** and **api**.</span></span>

1. <span data-ttu-id="43b43-123">Создайте файл с именем **auth.ts** в **каталоге ./src/api** и добавьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="43b43-123">Create a new file named **auth.ts** in the **./src/api** directory and add the following code.</span></span>

    ```typescript
    import Router from 'express-promise-router';

    const authRouter = Router();

    // TODO: Implement this router

    export default authRouter;
    ```

1. <span data-ttu-id="43b43-124">Создайте файл **graph.ts** в **каталоге ./src/api** и добавьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="43b43-124">Create a new file named **graph.ts** in the **./src/api** directory and add the following code.</span></span>

    ```typescript
    import Router from 'express-promise-router';

    const graphRouter = Router();

    // TODO: Implement this router

    export default graphRouter;
    ```

1. <span data-ttu-id="43b43-125">Создайте файл с **именем server.ts** в **каталоге ./src** и добавьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="43b43-125">Create a new file named **server.ts** in the **./src** directory and add the following code.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/server.ts" id="ServerSnippet":::

## <a name="create-the-add-in"></a><span data-ttu-id="43b43-126">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="43b43-126">Create the add-in</span></span>

1. <span data-ttu-id="43b43-127">Создайте файл с **именемtaskpane.html** в **каталоге ./src/addin** и добавьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="43b43-127">Create a new file named **taskpane.html** in the **./src/addin** directory and add the following code.</span></span>

    :::code language="html" source="../demo/graph-tutorial/src/addin/taskpane.html" id="TaskPaneHtmlSnippet":::

1. <span data-ttu-id="43b43-128">Создайте файл **taskpane.css** в **каталоге ./src/addin** и добавьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="43b43-128">Create a new file named **taskpane.css** in the **./src/addin** directory and add the following code.</span></span>

    :::code language="css" source="../demo/graph-tutorial/src/addin/taskpane.css":::

1. <span data-ttu-id="43b43-129">Создайте файл с **именемtaskpane.js** в **каталоге ./src/addin** и добавьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="43b43-129">Create a new file named **taskpane.js** in the **./src/addin** directory and add the following code.</span></span>

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

1. <span data-ttu-id="43b43-130">Создайте новый каталог в каталоге **.src/addin** с именем **assets.**</span><span class="sxs-lookup"><span data-stu-id="43b43-130">Create a new directory in the **.src/addin** directory named **assets**.</span></span>

1. <span data-ttu-id="43b43-131">Добавьте три файла PNG в этот каталог согласно следующей таблице.</span><span class="sxs-lookup"><span data-stu-id="43b43-131">Add three PNG files in this directory according to the following table.</span></span>

    | <span data-ttu-id="43b43-132">Имя файла</span><span class="sxs-lookup"><span data-stu-id="43b43-132">File name</span></span>   | <span data-ttu-id="43b43-133">Размер в пикселях</span><span class="sxs-lookup"><span data-stu-id="43b43-133">Size in pixels</span></span> |
    |-------------|----------------|
    | <span data-ttu-id="43b43-134">icon-80.png</span><span class="sxs-lookup"><span data-stu-id="43b43-134">icon-80.png</span></span> | <span data-ttu-id="43b43-135">80x80</span><span class="sxs-lookup"><span data-stu-id="43b43-135">80x80</span></span>          |
    | <span data-ttu-id="43b43-136">icon-32.png</span><span class="sxs-lookup"><span data-stu-id="43b43-136">icon-32.png</span></span> | <span data-ttu-id="43b43-137">32x32</span><span class="sxs-lookup"><span data-stu-id="43b43-137">32x32</span></span>          |
    | <span data-ttu-id="43b43-138">icon-16.png</span><span class="sxs-lookup"><span data-stu-id="43b43-138">icon-16.png</span></span> | <span data-ttu-id="43b43-139">16 x 16</span><span class="sxs-lookup"><span data-stu-id="43b43-139">16x16</span></span>          |

    > [!NOTE]
    > <span data-ttu-id="43b43-140">Для этого шага можно использовать любое изображение.</span><span class="sxs-lookup"><span data-stu-id="43b43-140">You can use any image you want for this step.</span></span> <span data-ttu-id="43b43-141">Вы также можете скачать изображения, используемые в этом примере, непосредственно с [GitHub.](https://github.com/microsoftgraph/msgraph-training-office-addin/demo/graph-tutorial/src/addin/assets)</span><span class="sxs-lookup"><span data-stu-id="43b43-141">You can also download the images used in this sample directly from [GitHub](https://github.com/microsoftgraph/msgraph-training-office-addin/demo/graph-tutorial/src/addin/assets).</span></span>

1. <span data-ttu-id="43b43-142">Создайте каталог в корневом каталоге **манифеста** проекта с именем .</span><span class="sxs-lookup"><span data-stu-id="43b43-142">Create a new directory in the root of the project named **manifest**.</span></span>

1. <span data-ttu-id="43b43-143">Создайте файл с **именемmanifest.xml** в **папке ./manifest** и добавьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="43b43-143">Create a new file named **manifest.xml** in the **./manifest** folder and add the following code.</span></span> <span data-ttu-id="43b43-144">Замените `NEW_GUID_HERE` новым GUID, например `b4fa03b8-1eb6-4e8b-a380-e0476be9e019` .</span><span class="sxs-lookup"><span data-stu-id="43b43-144">Replace `NEW_GUID_HERE` with a new GUID, like `b4fa03b8-1eb6-4e8b-a380-e0476be9e019`.</span></span>

    :::code language="xml" source="../demo/graph-tutorial/manifest/manifest.xml":::

## <a name="side-load-the-add-in-in-excel"></a><span data-ttu-id="43b43-145">Загрузка неогрузки надстройки в Excel</span><span class="sxs-lookup"><span data-stu-id="43b43-145">Side-load the add-in in Excel</span></span>

1. <span data-ttu-id="43b43-146">Запустите сервер с помощью следующей команды.</span><span class="sxs-lookup"><span data-stu-id="43b43-146">Start the server by running the following command.</span></span>

    ```Shell
    yarn start
    ```

1. <span data-ttu-id="43b43-147">Откройте браузер и перейдите к `https://localhost:3000/taskpane.html` .</span><span class="sxs-lookup"><span data-stu-id="43b43-147">Open your browser and browse to `https://localhost:3000/taskpane.html`.</span></span> <span data-ttu-id="43b43-148">Должно отобраться `Not loaded` сообщение.</span><span class="sxs-lookup"><span data-stu-id="43b43-148">You should see a `Not loaded` message.</span></span>

1. <span data-ttu-id="43b43-149">В браузере перейдите в [Office.com](https://www.office.com/) и войдите.</span><span class="sxs-lookup"><span data-stu-id="43b43-149">In your browser, go to [Office.com](https://www.office.com/) and sign in.</span></span> <span data-ttu-id="43b43-150">Выберите **"Создать"** на левой панели инструментов, а затем выберите **"Таблица".**</span><span class="sxs-lookup"><span data-stu-id="43b43-150">Select **Create** in the left-hand toolbar, then select **Spreadsheet**.</span></span>

    ![Снимок экрана: меню "Создать" Office.com](images/office-select-excel.png)

1. <span data-ttu-id="43b43-152">Выберите **вкладку "Вставка"** и выберите **надстройки Office.**</span><span class="sxs-lookup"><span data-stu-id="43b43-152">Select the **Insert** tab, then select **Office Add-ins**.</span></span>

1. <span data-ttu-id="43b43-153">Выберите **"Отправить мою надстройка"** и выберите **"Обзор".**</span><span class="sxs-lookup"><span data-stu-id="43b43-153">Select **Upload My Add-in**, then select **Browse**.</span></span> <span data-ttu-id="43b43-154">Загрузите **файл ./manifest/manifest.xml.**</span><span class="sxs-lookup"><span data-stu-id="43b43-154">Upload your **./manifest/manifest.xml** file.</span></span>

1. <span data-ttu-id="43b43-155">Выберите **кнопку "Импорт календаря"** на **вкладке "Главная",** чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="43b43-155">Select the **Import Calendar** button on the **Home** tab to open the taskpane.</span></span>

    ![Снимок экрана с кнопкой "Импорт календаря" на вкладке "Главная"](images/get-started.png)

1. <span data-ttu-id="43b43-157">После открытия области задач должно отобраться `Hello World!` сообщение.</span><span class="sxs-lookup"><span data-stu-id="43b43-157">After the taskpane opens, you should see a `Hello World!` message.</span></span>

    ![Снимок экрана: сообщение Hello World](images/hello-world.png)
