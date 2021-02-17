---
ms.openlocfilehash: 69d77c19cb2c589086df2b4bf19eeadf41aad58a
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274392"
---
<!-- markdownlint-disable MD002 MD041 -->

В этом упражнении вы встроим единый вход для надстройки [Office](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) и расширим веб-API для поддержки потока "от [имени".](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) Это необходимо для получения необходимого маркера доступа OAuth для вызова Microsoft Graph.

## <a name="overview"></a>Обзор

Служба SSO надстройки Office предоставляет маркер доступа, но он позволяет надстройки вызывать только собственный веб-API. Он не включает прямой доступ к Microsoft Graph. Процесс работает следующим образом.

1. Надстройка получает маркер, вызывая [getAccessToken.](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.auth?view=common-js#getaccesstoken-options-) Аудитория этого маркера (утверждение) — это ИД приложения для регистрации `aud` приложения надстройки.
1. Надстройка отправляет этот маркер в `Authorization` заголовке при вызове веб-API.
1. Веб-API проверяет маркер, а затем использует поток "от имени", чтобы обменять этот маркер на маркер Microsoft Graph. Аудитория нового маркера : `https://graph.microsoft.com` .
1. Веб-API использует новый маркер для вызова Microsoft Graph и возвращает результаты обратно надстройки.

## <a name="configure-the-solution"></a>Настройка решения

1. Откройте **./.env и** обновите ИД приложения и секрет клиента `AZURE_APP_ID` из `AZURE_CLIENT_SECRET` регистрации приложения.

    > [!IMPORTANT]
    > Если вы используете управление исходным кодом, например git, пришло бы время исключить **ENV-файл** из системы управления исходным кодом, чтобы не допустить случайной утечки вашего ИД приложения и секрета клиента.

1. Откройте **./manifest/manifest.xml** и замените все экземпляры на ИД приложения `YOUR_APP_ID_HERE` из регистрации приложения.

1. Создайте новый файл в **каталоге ./src/addin** с именем **config.js** и добавьте следующий код, заменив его на ИД приложения при `YOUR_APP_ID_HERE` регистрации приложения.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/config.example.js":::

## <a name="implement-sign-in"></a>Реализация входов

1. Откройте **файл ./src/api/auth.ts** и добавьте следующие утверждения в `import` верхней части файла.

    ```typescript
    import jwt, { SigningKeyCallback, JwtHeader } from 'jsonwebtoken';
    import jwksClient from 'jwks-rsa';
    import * as msal from '@azure/msal-node';
    ```

1. Добавьте следующий код после `import` заявок.

    :::code language="typescript" source="../demo/graph-tutorial/src/api/auth.ts" id="TokenExchangeSnippet":::

    Этот код [инициализирует конфиденциальный](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/initialize-confidential-client-application.md)клиент MSAL и экспортирует функцию для получения маркера Graph из маркера, отправленного надстройкой.

1. Добавьте следующий код перед `export default authRouter;` строкой.

    :::code language="typescript" source="../demo/graph-tutorial/src/api/auth.ts" id="GetAuthStatusSnippet":::

    Этот код реализует API ( ), который проверяет, можно ли автоматически обмениваться маркером надстройки на `GET /auth/status` маркер Graph. Надстройка будет использовать этот API, чтобы определить, нужно ли представить пользователю интерактивный вход.

1. Откройте **файл ./src/addin/taskpane.js** и добавьте в файл следующий код.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="AuthUiSnippet":::

    Этот код добавляет функции для обновления пользовательского интерфейса и использования [Dialog API Office](https://docs.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins) для инициации интерактивного потока проверки подлинности.

1. Добавьте следующую функцию для реализации временного основного пользовательского интерфейса.

    ```javascript
    function showMainUi() {
      $('.container').empty();
      $('<p/>', {
        class: 'ms-fontSize-24 ms-fontWeight-bold',
        text: 'Authenticated!'
      }).appendTo('.container');
    }
    ```

1. Замените `Office.onReady` существующий вызов следующим:

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="OfficeReadySnippet":::

    Подумайте, что делает этот код.

    - При первой загрузке области задач она вызывает, чтобы получить маркер с областью действия для `getAccessToken` веб-API надстройки.
    - Он использует этот маркер для вызова API, чтобы проверить, предоставил ли пользователь согласие на использование областей `/auth/status` Microsoft Graph.
        - Если пользователь не согласился, он использует всплывающее окно для получения согласия пользователя с помощью интерактивного входа.
        - Если пользователь согласился, он загружает основной пользовательский интерфейс.

### <a name="getting-user-consent"></a>Получение согласия пользователя

Несмотря на то что надстройка использует службу SSO, пользователю все равно необходимо дать согласие на доступ надстройки к своим данным через Microsoft Graph. Получение согласия — это разовая процедура. После предоставления пользователем согласия маркер SSO можно обменять на маркер Graph без участия пользователя. В этом разделе вы реализуйте в надстройки согласие с помощью [msal-browser.](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser)

1. Создайте новый файл в **каталоге ./src/addin** с именемconsent.jsи **добавьте** следующий код.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/consent.js" id="ConsentJsSnippet":::

    Этот код делает вход для пользователя, запрашивая набор разрешений Microsoft Graph, настроенных при регистрации приложения.

1. Создайте новый файл в **каталоге ./src/addin** с именем **consent.html** и добавьте следующий код.

    :::code language="html" source="../demo/graph-tutorial/src/addin/consent.html" id="ConsentHtmlSnippet":::

    Этот код реализует базовую HTML-страницу для **загрузки** consent.jsфайла. Эта страница будет загружена во всплывающее диалоговое окно.

1. Сохраните все изменения и перезапустите сервер.

1. Повторно загрузите файл **manifest.xml,** выполнив те же действия, что и при загрузке надстройки в [Excel.](02-create-app.md#side-load-the-add-in-in-excel)

1. Выберите **кнопку "Импорт календаря"** на **вкладке "Главная",** чтобы открыть ее.

1. Выберите **кнопку "Предоставить разрешение"** в области задач, чтобы запустить диалоговое окно согласия во всплывающее окно. Во sign in and grant consent.

1. The task pane updates with an "Authenticated!" Сообщение. Вы можете проверить маркеры следующим образом.

    - В средствах разработчика brower маркер API отображается в консоли.
    - В CLI, где вы работаете на Node.js, печатается маркер Graph.

    Вы можете сравнить эти маркеры по [https://jwt.ms](https://jwt.ms) . Обратите внимание, что аудитория маркера API ( ) имеет ИД приложения регистрации приложения, а область `aud` ( `scp` ) — `access_as_user` .
