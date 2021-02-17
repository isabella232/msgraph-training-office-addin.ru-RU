---
ms.openlocfilehash: 69d77c19cb2c589086df2b4bf19eeadf41aad58a
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274392"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="a7236-101">В этом упражнении вы встроим единый вход для надстройки [Office](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) и расширим веб-API для поддержки потока "от [имени".](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)</span><span class="sxs-lookup"><span data-stu-id="a7236-101">In this exercise you will enable [Office Add-in single sign-on (SSO)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) in the add-in, and extend the web API to support [on-behalf-of flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow).</span></span> <span data-ttu-id="a7236-102">Это необходимо для получения необходимого маркера доступа OAuth для вызова Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a7236-102">This is required to obtain the necessary OAuth access token to call the Microsoft Graph.</span></span>

## <a name="overview"></a><span data-ttu-id="a7236-103">Обзор</span><span class="sxs-lookup"><span data-stu-id="a7236-103">Overview</span></span>

<span data-ttu-id="a7236-104">Служба SSO надстройки Office предоставляет маркер доступа, но он позволяет надстройки вызывать только собственный веб-API.</span><span class="sxs-lookup"><span data-stu-id="a7236-104">Office Add-in SSO provides an access token, but that token is only enables the add-in to call it's own web API.</span></span> <span data-ttu-id="a7236-105">Он не включает прямой доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a7236-105">It does not enable direct access to the Microsoft Graph.</span></span> <span data-ttu-id="a7236-106">Процесс работает следующим образом.</span><span class="sxs-lookup"><span data-stu-id="a7236-106">The process works as follows.</span></span>

1. <span data-ttu-id="a7236-107">Надстройка получает маркер, вызывая [getAccessToken.](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.auth?view=common-js#getaccesstoken-options-)</span><span class="sxs-lookup"><span data-stu-id="a7236-107">The add-in gets a token by calling [getAccessToken](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.auth?view=common-js#getaccesstoken-options-).</span></span> <span data-ttu-id="a7236-108">Аудитория этого маркера (утверждение) — это ИД приложения для регистрации `aud` приложения надстройки.</span><span class="sxs-lookup"><span data-stu-id="a7236-108">This token's audience (the `aud` claim) is the application ID of the add-in's app registration.</span></span>
1. <span data-ttu-id="a7236-109">Надстройка отправляет этот маркер в `Authorization` заголовке при вызове веб-API.</span><span class="sxs-lookup"><span data-stu-id="a7236-109">The add-in sends this token in the `Authorization` header when it makes a call to the web API.</span></span>
1. <span data-ttu-id="a7236-110">Веб-API проверяет маркер, а затем использует поток "от имени", чтобы обменять этот маркер на маркер Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a7236-110">The web API validates the token, then uses the on-behalf-of flow to exchange this token for a Microsoft Graph token.</span></span> <span data-ttu-id="a7236-111">Аудитория нового маркера : `https://graph.microsoft.com` .</span><span class="sxs-lookup"><span data-stu-id="a7236-111">This new token's audience is `https://graph.microsoft.com`.</span></span>
1. <span data-ttu-id="a7236-112">Веб-API использует новый маркер для вызова Microsoft Graph и возвращает результаты обратно надстройки.</span><span class="sxs-lookup"><span data-stu-id="a7236-112">The web API uses the new token to make calls to the Microsoft Graph, and returns the results back to the add-in.</span></span>

## <a name="configure-the-solution"></a><span data-ttu-id="a7236-113">Настройка решения</span><span class="sxs-lookup"><span data-stu-id="a7236-113">Configure the solution</span></span>

1. <span data-ttu-id="a7236-114">Откройте **./.env и** обновите ИД приложения и секрет клиента `AZURE_APP_ID` из `AZURE_CLIENT_SECRET` регистрации приложения.</span><span class="sxs-lookup"><span data-stu-id="a7236-114">Open **./.env** and update the `AZURE_APP_ID` and `AZURE_CLIENT_SECRET` with the application ID and client secret from your app registration.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="a7236-115">Если вы используете управление исходным кодом, например git, пришло бы время исключить **ENV-файл** из системы управления исходным кодом, чтобы не допустить случайной утечки вашего ИД приложения и секрета клиента.</span><span class="sxs-lookup"><span data-stu-id="a7236-115">If you're using source control such as git, now would be a good time to exclude the **.env** file from source control to avoid inadvertently leaking your app ID and client secret.</span></span>

1. <span data-ttu-id="a7236-116">Откройте **./manifest/manifest.xml** и замените все экземпляры на ИД приложения `YOUR_APP_ID_HERE` из регистрации приложения.</span><span class="sxs-lookup"><span data-stu-id="a7236-116">Open **./manifest/manifest.xml** and replace all instances of `YOUR_APP_ID_HERE` with the application ID from your app registration.</span></span>

1. <span data-ttu-id="a7236-117">Создайте новый файл в **каталоге ./src/addin** с именем **config.js** и добавьте следующий код, заменив его на ИД приложения при `YOUR_APP_ID_HERE` регистрации приложения.</span><span class="sxs-lookup"><span data-stu-id="a7236-117">Create a new file in the **./src/addin** directory named **config.js** and add the following code, replacing `YOUR_APP_ID_HERE` with the application ID from your app registration.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/config.example.js":::

## <a name="implement-sign-in"></a><span data-ttu-id="a7236-118">Реализация входов</span><span class="sxs-lookup"><span data-stu-id="a7236-118">Implement sign-in</span></span>

1. <span data-ttu-id="a7236-119">Откройте **файл ./src/api/auth.ts** и добавьте следующие утверждения в `import` верхней части файла.</span><span class="sxs-lookup"><span data-stu-id="a7236-119">Open **./src/api/auth.ts** and add the following `import` statements at the top of the file.</span></span>

    ```typescript
    import jwt, { SigningKeyCallback, JwtHeader } from 'jsonwebtoken';
    import jwksClient from 'jwks-rsa';
    import * as msal from '@azure/msal-node';
    ```

1. <span data-ttu-id="a7236-120">Добавьте следующий код после `import` заявок.</span><span class="sxs-lookup"><span data-stu-id="a7236-120">Add the following code after the `import` statements.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/auth.ts" id="TokenExchangeSnippet":::

    <span data-ttu-id="a7236-121">Этот код [инициализирует конфиденциальный](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/initialize-confidential-client-application.md)клиент MSAL и экспортирует функцию для получения маркера Graph из маркера, отправленного надстройкой.</span><span class="sxs-lookup"><span data-stu-id="a7236-121">This code [initializes an MSAL confidential client](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/initialize-confidential-client-application.md), and exports a function to get a Graph token from the token sent by the add-in.</span></span>

1. <span data-ttu-id="a7236-122">Добавьте следующий код перед `export default authRouter;` строкой.</span><span class="sxs-lookup"><span data-stu-id="a7236-122">Add the following code before the `export default authRouter;` line.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/auth.ts" id="GetAuthStatusSnippet":::

    <span data-ttu-id="a7236-123">Этот код реализует API ( ), который проверяет, можно ли автоматически обмениваться маркером надстройки на `GET /auth/status` маркер Graph.</span><span class="sxs-lookup"><span data-stu-id="a7236-123">This code implements an API (`GET /auth/status`) that checks if the add-in token can be silently exchanged for a Graph token.</span></span> <span data-ttu-id="a7236-124">Надстройка будет использовать этот API, чтобы определить, нужно ли представить пользователю интерактивный вход.</span><span class="sxs-lookup"><span data-stu-id="a7236-124">The add-in will use this API to determine if it needs to present an interactive login to the user.</span></span>

1. <span data-ttu-id="a7236-125">Откройте **файл ./src/addin/taskpane.js** и добавьте в файл следующий код.</span><span class="sxs-lookup"><span data-stu-id="a7236-125">Open **./src/addin/taskpane.js** and add the following code to the file.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="AuthUiSnippet":::

    <span data-ttu-id="a7236-126">Этот код добавляет функции для обновления пользовательского интерфейса и использования [Dialog API Office](https://docs.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins) для инициации интерактивного потока проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="a7236-126">This code adds functions to update the UI, and to use the [Office Dialog API](https://docs.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins) to initiate an interactive authentication flow.</span></span>

1. <span data-ttu-id="a7236-127">Добавьте следующую функцию для реализации временного основного пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="a7236-127">Add the following function to implement a temporary main UI.</span></span>

    ```javascript
    function showMainUi() {
      $('.container').empty();
      $('<p/>', {
        class: 'ms-fontSize-24 ms-fontWeight-bold',
        text: 'Authenticated!'
      }).appendTo('.container');
    }
    ```

1. <span data-ttu-id="a7236-128">Замените `Office.onReady` существующий вызов следующим:</span><span class="sxs-lookup"><span data-stu-id="a7236-128">Replace the existing `Office.onReady` call with the following.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="OfficeReadySnippet":::

    <span data-ttu-id="a7236-129">Подумайте, что делает этот код.</span><span class="sxs-lookup"><span data-stu-id="a7236-129">Consider what this code does.</span></span>

    - <span data-ttu-id="a7236-130">При первой загрузке области задач она вызывает, чтобы получить маркер с областью действия для `getAccessToken` веб-API надстройки.</span><span class="sxs-lookup"><span data-stu-id="a7236-130">When the task pane first loads, it calls `getAccessToken` to get a token scoped for the add-in's web API.</span></span>
    - <span data-ttu-id="a7236-131">Он использует этот маркер для вызова API, чтобы проверить, предоставил ли пользователь согласие на использование областей `/auth/status` Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a7236-131">It uses that token to call the `/auth/status` API to check if the user has given consent to the Microsoft Graph scopes yet.</span></span>
        - <span data-ttu-id="a7236-132">Если пользователь не согласился, он использует всплывающее окно для получения согласия пользователя с помощью интерактивного входа.</span><span class="sxs-lookup"><span data-stu-id="a7236-132">If the user has not consented, it uses a pop-up window to get the user's consent through an interactive login.</span></span>
        - <span data-ttu-id="a7236-133">Если пользователь согласился, он загружает основной пользовательский интерфейс.</span><span class="sxs-lookup"><span data-stu-id="a7236-133">If the user has consented, it loads the main UI.</span></span>

### <a name="getting-user-consent"></a><span data-ttu-id="a7236-134">Получение согласия пользователя</span><span class="sxs-lookup"><span data-stu-id="a7236-134">Getting user consent</span></span>

<span data-ttu-id="a7236-135">Несмотря на то что надстройка использует службу SSO, пользователю все равно необходимо дать согласие на доступ надстройки к своим данным через Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a7236-135">Even though the add-in is using SSO, the user still has to consent to the add-in accessing their data via Microsoft Graph.</span></span> <span data-ttu-id="a7236-136">Получение согласия — это разовая процедура.</span><span class="sxs-lookup"><span data-stu-id="a7236-136">Getting consent is a one-time process.</span></span> <span data-ttu-id="a7236-137">После предоставления пользователем согласия маркер SSO можно обменять на маркер Graph без участия пользователя.</span><span class="sxs-lookup"><span data-stu-id="a7236-137">Once the user has granted consent, the SSO token can be exchanged for a Graph token without any user interaction.</span></span> <span data-ttu-id="a7236-138">В этом разделе вы реализуйте в надстройки согласие с помощью [msal-browser.](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser)</span><span class="sxs-lookup"><span data-stu-id="a7236-138">In this section you'll implement the consent experience in the add-in using [msal-browser](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser).</span></span>

1. <span data-ttu-id="a7236-139">Создайте новый файл в **каталоге ./src/addin** с именемconsent.jsи **добавьте** следующий код.</span><span class="sxs-lookup"><span data-stu-id="a7236-139">Create a new file in the **./src/addin** directory named **consent.js** and add the following code.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/consent.js" id="ConsentJsSnippet":::

    <span data-ttu-id="a7236-140">Этот код делает вход для пользователя, запрашивая набор разрешений Microsoft Graph, настроенных при регистрации приложения.</span><span class="sxs-lookup"><span data-stu-id="a7236-140">This code does login for the user, requesting the set of Microsoft Graph permissions that are configured on the app registration.</span></span>

1. <span data-ttu-id="a7236-141">Создайте новый файл в **каталоге ./src/addin** с именем **consent.html** и добавьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="a7236-141">Create a new file in the **./src/addin** directory named **consent.html** and add the following code.</span></span>

    :::code language="html" source="../demo/graph-tutorial/src/addin/consent.html" id="ConsentHtmlSnippet":::

    <span data-ttu-id="a7236-142">Этот код реализует базовую HTML-страницу для **загрузки** consent.jsфайла.</span><span class="sxs-lookup"><span data-stu-id="a7236-142">This code implements a basic HTML page to load the **consent.js** file.</span></span> <span data-ttu-id="a7236-143">Эта страница будет загружена во всплывающее диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="a7236-143">This page will be loaded in a pop-up dialog.</span></span>

1. <span data-ttu-id="a7236-144">Сохраните все изменения и перезапустите сервер.</span><span class="sxs-lookup"><span data-stu-id="a7236-144">Save all of your changes and restart the server.</span></span>

1. <span data-ttu-id="a7236-145">Повторно загрузите файл **manifest.xml,** выполнив те же действия, что и при загрузке надстройки в [Excel.](02-create-app.md#side-load-the-add-in-in-excel)</span><span class="sxs-lookup"><span data-stu-id="a7236-145">Re-upload your **manifest.xml** file using the same steps in [Side-load the add-in in Excel](02-create-app.md#side-load-the-add-in-in-excel).</span></span>

1. <span data-ttu-id="a7236-146">Выберите **кнопку "Импорт календаря"** на **вкладке "Главная",** чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="a7236-146">Select the **Import Calendar** button on the **Home** tab to open the task pane.</span></span>

1. <span data-ttu-id="a7236-147">Выберите **кнопку "Предоставить разрешение"** в области задач, чтобы запустить диалоговое окно согласия во всплывающее окно.</span><span class="sxs-lookup"><span data-stu-id="a7236-147">Select the **Give permission** button in the task pane to launch the consent dialog in a pop-up window.</span></span> <span data-ttu-id="a7236-148">Во sign in and grant consent.</span><span class="sxs-lookup"><span data-stu-id="a7236-148">Sign in and grant consent.</span></span>

1. <span data-ttu-id="a7236-149">The task pane updates with an "Authenticated!"</span><span class="sxs-lookup"><span data-stu-id="a7236-149">The task pane updates with an "Authenticated!"</span></span> <span data-ttu-id="a7236-150">Сообщение.</span><span class="sxs-lookup"><span data-stu-id="a7236-150">message.</span></span> <span data-ttu-id="a7236-151">Вы можете проверить маркеры следующим образом.</span><span class="sxs-lookup"><span data-stu-id="a7236-151">You can check the tokens as follows.</span></span>

    - <span data-ttu-id="a7236-152">В средствах разработчика brower маркер API отображается в консоли.</span><span class="sxs-lookup"><span data-stu-id="a7236-152">In your brower's developer tools, the API token is shown in the Console.</span></span>
    - <span data-ttu-id="a7236-153">В CLI, где вы работаете на Node.js, печатается маркер Graph.</span><span class="sxs-lookup"><span data-stu-id="a7236-153">In your CLI where you are running the Node.js server, the Graph token is printed.</span></span>

    <span data-ttu-id="a7236-154">Вы можете сравнить эти маркеры по [https://jwt.ms](https://jwt.ms) .</span><span class="sxs-lookup"><span data-stu-id="a7236-154">You can compare these token at [https://jwt.ms](https://jwt.ms).</span></span> <span data-ttu-id="a7236-155">Обратите внимание, что аудитория маркера API ( ) имеет ИД приложения регистрации приложения, а область `aud` ( `scp` ) — `access_as_user` .</span><span class="sxs-lookup"><span data-stu-id="a7236-155">Notice that the API token's audience (`aud`) is set to the application ID of your app registration, and the scope (`scp`) is `access_as_user`.</span></span>
