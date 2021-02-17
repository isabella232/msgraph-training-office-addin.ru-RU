---
ms.openlocfilehash: fed56663591ac36e4defcae3a8d0a222f86e3026
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274370"
---
# <a name="how-to-run-the-completed-project"></a><span data-ttu-id="92134-101">Как запустить завершенный проект</span><span class="sxs-lookup"><span data-stu-id="92134-101">How to run the completed project</span></span>

## <a name="prerequisites"></a><span data-ttu-id="92134-102">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="92134-102">Prerequisites</span></span>

<span data-ttu-id="92134-103">Чтобы запустить завершенный проект в этой папке, необходимо следующее:</span><span class="sxs-lookup"><span data-stu-id="92134-103">To run the completed project in this folder, you need the following:</span></span>

- <span data-ttu-id="92134-104">[Node.js](https://nodejs.org) [и Yarn,](https://yarnpkg.com/) установленные на компьютере для разработки.</span><span class="sxs-lookup"><span data-stu-id="92134-104">[Node.js](https://nodejs.org) and [Yarn](https://yarnpkg.com/) installed on your development machine.</span></span> <span data-ttu-id="92134-105">(**Примечание.** Это руководство было написано на node версии 14.15.0 и Yarn версии 1.22.0.</span><span class="sxs-lookup"><span data-stu-id="92134-105">(**Note:** This tutorial was written with Node version 14.15.0 and Yarn version 1.22.0.</span></span> <span data-ttu-id="92134-106">Действия в этом руководстве могут работать с другими версиями, но не были протестированы.)</span><span class="sxs-lookup"><span data-stu-id="92134-106">The steps in this guide may work with other versions, but that has not been tested.)</span></span>
- <span data-ttu-id="92134-107">Личная учетная запись Майкрософт с почтовым ящиком Outlook.com или учетная запись Майкрософт для работы или учебного заведения.</span><span class="sxs-lookup"><span data-stu-id="92134-107">Either a personal Microsoft account with a mailbox on Outlook.com, or a Microsoft work or school account.</span></span>

<span data-ttu-id="92134-108">Если у вас нет учетной записи Майкрософт, существует несколько вариантов получения бесплатной учетной записи:</span><span class="sxs-lookup"><span data-stu-id="92134-108">If you don't have a Microsoft account, there are a couple of options to get a free account:</span></span>

- <span data-ttu-id="92134-109">Вы можете [зарегистрироваться для новой личной учетной записи Майкрософт.](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1)</span><span class="sxs-lookup"><span data-stu-id="92134-109">You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).</span></span>
- <span data-ttu-id="92134-110">Вы можете зарегистрироваться в программе для разработчиков [Office 365,](https://developer.microsoft.com/office/dev-program) чтобы получить бесплатную подписку на Office 365.</span><span class="sxs-lookup"><span data-stu-id="92134-110">You can [sign up for the Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get a free Office 365 subscription.</span></span>

## <a name="register-a-web-application-with-the-azure-active-directory-admin-center"></a><span data-ttu-id="92134-111">Регистрация веб-приложения в Центре администрирования Azure Active Directory</span><span class="sxs-lookup"><span data-stu-id="92134-111">Register a web application with the Azure Active Directory admin center</span></span>

1. <span data-ttu-id="92134-112">Откройте браузер и перейдите к [Центру администрирования Azure Active Directory](https://aad.portal.azure.com).</span><span class="sxs-lookup"><span data-stu-id="92134-112">Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com).</span></span> <span data-ttu-id="92134-113">Войдите с помощью **личной учетной записи** (т.е. учетной записи Microsoft) или **рабочей (учебной) учетной записи**.</span><span class="sxs-lookup"><span data-stu-id="92134-113">Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.</span></span>

1. <span data-ttu-id="92134-114">Выберите **Azure Active Directory** на панели навигации слева, затем выберите **Регистрация приложений** в разделе **Управление**.</span><span class="sxs-lookup"><span data-stu-id="92134-114">Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.</span></span>

    ![<span data-ttu-id="92134-115">Снимок экрана с регистрацией приложений</span><span class="sxs-lookup"><span data-stu-id="92134-115">A screenshot of the App registrations</span></span> ](/tutorial/images/app-registrations.png)

1. <span data-ttu-id="92134-116">Выберите **Новая регистрация**.</span><span class="sxs-lookup"><span data-stu-id="92134-116">Select **New registration**.</span></span> <span data-ttu-id="92134-117">На странице **Зарегистрировать приложение** задайте необходимые значения следующим образом.</span><span class="sxs-lookup"><span data-stu-id="92134-117">On the **Register an application** page, set the values as follows.</span></span>

    - <span data-ttu-id="92134-118">Введите **имя** `Office Add-in Graph Tutorial`.</span><span class="sxs-lookup"><span data-stu-id="92134-118">Set **Name** to `Office Add-in Graph Tutorial`.</span></span>
    - <span data-ttu-id="92134-119">Введите **поддерживаемые типы учетных записей** для **учетных записей в любом каталоге организаций и личных учетных записей Microsoft**.</span><span class="sxs-lookup"><span data-stu-id="92134-119">Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.</span></span>
    - <span data-ttu-id="92134-120">В разделе **URI адрес перенаправления** введите значение в первом раскрывающемся списке `Single-page application (SPA)` и задайте значение `https://localhost:3000/consent.html`.</span><span class="sxs-lookup"><span data-stu-id="92134-120">Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://localhost:3000/consent.html`.</span></span>

    ![Снимок экрана: страница "Регистрация приложения"](/tutorial/images/register-an-app.png)

1. <span data-ttu-id="92134-122">Нажмите **Зарегистрировать**.</span><span class="sxs-lookup"><span data-stu-id="92134-122">Select **Register**.</span></span> <span data-ttu-id="92134-123">На странице **учебника по Graph** для надстройки Office скопируйте значение ИД приложения **(клиента)** и сохраните его, оно потребуется вам на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="92134-123">On the **Office Add-in Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.</span></span>

    ![Снимок экрана: ИД нового приложения для регистрации](/tutorial/images/application-id.png)

1. <span data-ttu-id="92134-125">Выберите **Сертификаты и секреты** в разделе **Управление**.</span><span class="sxs-lookup"><span data-stu-id="92134-125">Select **Certificates & secrets** under **Manage**.</span></span> <span data-ttu-id="92134-126">Нажмите кнопку **Новый секрет клиента**.</span><span class="sxs-lookup"><span data-stu-id="92134-126">Select the **New client secret** button.</span></span> <span data-ttu-id="92134-127">Введите значение в **описании** и выберите один из параметров **"Срок действия истекает"** и выберите **"Добавить".**</span><span class="sxs-lookup"><span data-stu-id="92134-127">Enter a value in **Description** and select one of the options for **Expires** and select **Add**.</span></span>

1. <span data-ttu-id="92134-128">Скопируйте значение секрета клиента, а затем покиньте эту страницу.</span><span class="sxs-lookup"><span data-stu-id="92134-128">Copy the client secret value before you leave this page.</span></span> <span data-ttu-id="92134-129">Оно вам понадобится на следующем шаге.</span><span class="sxs-lookup"><span data-stu-id="92134-129">You will need it in the next step.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="92134-130">Это секрет клиента, он никогда не отображается еще раз, поэтому убедитесь, что вы скопировали его.</span><span class="sxs-lookup"><span data-stu-id="92134-130">This client secret is never shown again, so make sure you copy it now.</span></span>

1. <span data-ttu-id="92134-131">Выберите **разрешения API в** области **"Управление",** а затем **выберите "Добавить разрешение".**</span><span class="sxs-lookup"><span data-stu-id="92134-131">Select **API permissions** under **Manage**, then select **Add a permission**.</span></span>

1. <span data-ttu-id="92134-132">Выберите **Microsoft Graph,** а затем **делегирование разрешений.**</span><span class="sxs-lookup"><span data-stu-id="92134-132">Select **Microsoft Graph**, then **Delegated permissions**.</span></span>

1. <span data-ttu-id="92134-133">Выберите следующие разрешения, а затем выберите **"Добавить разрешения".**</span><span class="sxs-lookup"><span data-stu-id="92134-133">Select the following permissions, then select **Add permissions**.</span></span>

    - <span data-ttu-id="92134-134">**Calendars.ReadWrite** — позволяет приложению читать и записывать данные в календарь пользователя.</span><span class="sxs-lookup"><span data-stu-id="92134-134">**Calendars.ReadWrite** - this will allow the app to read and write to the user's calendar.</span></span>
    - <span data-ttu-id="92134-135">**MailboxSettings.Read** — позволяет приложению получить часовой пояс пользователя из параметров почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="92134-135">**MailboxSettings.Read** - this will allow the app to get the user's time zone from their mailbox settings.</span></span>

    ![Снимок экрана с настроенными разрешениями](/tutorial/images/configured-permissions.png)

## <a name="configure-office-add-in-single-sign-on"></a><span data-ttu-id="92134-137">Настройка единого входов для надстройки Office</span><span class="sxs-lookup"><span data-stu-id="92134-137">Configure Office Add-in single sign-on</span></span>

<span data-ttu-id="92134-138">Обновив регистрацию приложения для поддержки единого входов [(SSO) надстройки Office.](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)</span><span class="sxs-lookup"><span data-stu-id="92134-138">Update the app registration to support [Office Add-in single sign-on (SSO)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).</span></span>

1. <span data-ttu-id="92134-139">Выберите **"Показать API"**.</span><span class="sxs-lookup"><span data-stu-id="92134-139">Select **Expose an API**.</span></span> <span data-ttu-id="92134-140">В разделе **"Области", определяемом этим API,** выберите **"Добавить область".**</span><span class="sxs-lookup"><span data-stu-id="92134-140">In the **Scopes defined by this API** section, select **Add a scope**.</span></span> <span data-ttu-id="92134-141">Когда вам будет предложено установить **URI** для ИД приложения, установите значение , заменив его `api://localhost:3000/YOUR_APP_ID_HERE` на `YOUR_APP_ID_HERE` ИД приложения.</span><span class="sxs-lookup"><span data-stu-id="92134-141">When prompted to set an **Application ID URI**, set the value to `api://localhost:3000/YOUR_APP_ID_HERE`, replacing `YOUR_APP_ID_HERE` with the application ID.</span></span> <span data-ttu-id="92134-142">Choose **Save and continue**.</span><span class="sxs-lookup"><span data-stu-id="92134-142">Choose **Save and continue**.</span></span>

1. <span data-ttu-id="92134-143">Заполните поля следующим образом и выберите **"Добавить область".**</span><span class="sxs-lookup"><span data-stu-id="92134-143">Fill in the fields as follows and select **Add scope**.</span></span>

    - <span data-ttu-id="92134-144">**Имя области:**`access_as_user`</span><span class="sxs-lookup"><span data-stu-id="92134-144">**Scope name:** `access_as_user`</span></span>
    - <span data-ttu-id="92134-145">**Кто может дать согласие?: администраторы и пользователи**</span><span class="sxs-lookup"><span data-stu-id="92134-145">**Who can consent?: Admins and users**</span></span>
    - <span data-ttu-id="92134-146">**Отображаемая имя согласия администратора:**`Access the app as the user`</span><span class="sxs-lookup"><span data-stu-id="92134-146">**Admin consent display name:** `Access the app as the user`</span></span>
    - <span data-ttu-id="92134-147">**Описание согласия администратора:**`Allows Office Add-ins to call the app's web APIs as the current user.`</span><span class="sxs-lookup"><span data-stu-id="92134-147">**Admin consent description:** `Allows Office Add-ins to call the app's web APIs as the current user.`</span></span>
    - <span data-ttu-id="92134-148">**Отображаемая имя согласия пользователя:**`Access the app as you`</span><span class="sxs-lookup"><span data-stu-id="92134-148">**User consent display name:** `Access the app as you`</span></span>
    - <span data-ttu-id="92134-149">**Описание согласия пользователя:**`Allows Office Add-ins to call the app's web APIs as you.`</span><span class="sxs-lookup"><span data-stu-id="92134-149">**User consent description:** `Allows Office Add-ins to call the app's web APIs as you.`</span></span>
    - <span data-ttu-id="92134-150">**Состояние: включено**</span><span class="sxs-lookup"><span data-stu-id="92134-150">**State: Enabled**</span></span>

    ![Снимок экрана с формой "Добавление области"](/tutorial/images/add-scope.png)

1. <span data-ttu-id="92134-152">В разделе **"Авторизованные клиентские приложения"** выберите **"Добавить клиентские приложения".**</span><span class="sxs-lookup"><span data-stu-id="92134-152">In the **Authorized client applications** section, select **Add a client application**.</span></span> <span data-ttu-id="92134-153">Введите ИД клиента из следующего списка, введите область в разрешенных разрешениях и выберите **"Добавить приложение".**</span><span class="sxs-lookup"><span data-stu-id="92134-153">Enter a client ID from the following list, enable the scope under **Authorized scopes**, and select **Add application**.</span></span> <span data-ttu-id="92134-154">Повторите этот процесс для каждого из клиентских ИД в списке.</span><span class="sxs-lookup"><span data-stu-id="92134-154">Repeat this process for each of the client IDs in the list.</span></span>

    - <span data-ttu-id="92134-155">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="92134-155">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="92134-156">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="92134-156">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="92134-157">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="92134-157">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="92134-158">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="92134-158">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span></span>

## <a name="install-development-certificates"></a><span data-ttu-id="92134-159">Установка сертификатов разработки</span><span class="sxs-lookup"><span data-stu-id="92134-159">Install development certificates</span></span>

1. <span data-ttu-id="92134-160">Чтобы создать и установить сертификаты разработки для надстройки, запустите следующую команду.</span><span class="sxs-lookup"><span data-stu-id="92134-160">Run the following command to generate and install development certificates for your add-in.</span></span>

    ```Shell
    npx office-addin-dev-certs install
    ```

    <span data-ttu-id="92134-161">Если вам будет предложено подтвердить, подтвердит действия.</span><span class="sxs-lookup"><span data-stu-id="92134-161">If prompted for confirmation, confirm the actions.</span></span> <span data-ttu-id="92134-162">После выполнения команды вы увидите выходные данные, похожие на следующие.</span><span class="sxs-lookup"><span data-stu-id="92134-162">Once the command completes, you will see output similar to the following.</span></span>

    ```Shell
    You now have trusted access to https://localhost.
    Certificate: <path>\localhost.crt
    Key: <path>\localhost.key
    ```

1. <span data-ttu-id="92134-163">Скопируйте пути к localhost.crt и localhost.key, они вам потребуется на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="92134-163">Copy the paths to localhost.crt and localhost.key, you'll need them in the next step.</span></span>

## <a name="update-the-manifest"></a><span data-ttu-id="92134-164">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="92134-164">Update the manifest</span></span>

1. <span data-ttu-id="92134-165">Откройте файл **manifest.xml** и внести следующие изменения.</span><span class="sxs-lookup"><span data-stu-id="92134-165">Open the **manifest.xml** file and make the following changes.</span></span>
    1. <span data-ttu-id="92134-166">Замените `NEW_GUID_HERE` новым GUID, например `b4fa03b8-1eb6-4e8b-a380-e0476be9e019` .</span><span class="sxs-lookup"><span data-stu-id="92134-166">Replace `NEW_GUID_HERE` with a new GUID, like `b4fa03b8-1eb6-4e8b-a380-e0476be9e019`.</span></span>
    1. <span data-ttu-id="92134-167">Замените все экземпляры `YOUR_APP_ID_HERE` на ИД приложения из регистрации приложения.</span><span class="sxs-lookup"><span data-stu-id="92134-167">Replace all instances of `YOUR_APP_ID_HERE` with the application ID from your app registration.</span></span>

## <a name="configure-the-sample"></a><span data-ttu-id="92134-168">Настройка примера</span><span class="sxs-lookup"><span data-stu-id="92134-168">Configure the sample</span></span>

1. <span data-ttu-id="92134-169">Переименуем `example.env` файл в `.env` .</span><span class="sxs-lookup"><span data-stu-id="92134-169">Rename the `example.env` file to `.env`.</span></span>
1. <span data-ttu-id="92134-170">`.env`Отредактируете файл и внести следующие изменения.</span><span class="sxs-lookup"><span data-stu-id="92134-170">Edit the `.env` file and make the following changes.</span></span>
    1. <span data-ttu-id="92134-171">`YOUR_APP_ID_HERE`Замените **ид приложения,** который вы получили с портала регистрации приложений.</span><span class="sxs-lookup"><span data-stu-id="92134-171">Replace `YOUR_APP_ID_HERE` with the **Application Id** you got from the App Registration Portal.</span></span>
    1. <span data-ttu-id="92134-172">Замените `YOUR_CLIENT_SECRET_HERE` секрет клиента, который вы получили с портала регистрации приложений.</span><span class="sxs-lookup"><span data-stu-id="92134-172">Replace `YOUR_CLIENT_SECRET_HERE` with the client secret you got from the App Registration Portal.</span></span>
    1. <span data-ttu-id="92134-173">Замените `PATH_TO_LOCALHOST.CRT` путь к файлу localhost.crt из выходных данных `npx office-addin-dev-certs install` команды.</span><span class="sxs-lookup"><span data-stu-id="92134-173">Replace `PATH_TO_LOCALHOST.CRT` with the path to your localhost.crt file from the output of the `npx office-addin-dev-certs install` command.</span></span>
    1. <span data-ttu-id="92134-174">Замените `PATH_TO_LOCALHOST.KEY` путь к файлу localhost.key из выходных данных `npx office-addin-dev-certs install` команды.</span><span class="sxs-lookup"><span data-stu-id="92134-174">Replace `PATH_TO_LOCALHOST.KEY` with the path to your localhost.key file from the output of the `npx office-addin-dev-certs install` command.</span></span>

1. <span data-ttu-id="92134-175">Переименуем `config.example.js` файл в `config.js` .</span><span class="sxs-lookup"><span data-stu-id="92134-175">Rename the `config.example.js` file to `config.js`.</span></span>
1. <span data-ttu-id="92134-176">`config.js`Отредактируете файл и внести следующие изменения.</span><span class="sxs-lookup"><span data-stu-id="92134-176">Edit the `config.js` file and make the following changes.</span></span>
    1. <span data-ttu-id="92134-177">`YOUR_APP_ID_HERE`Замените **ид приложения,** который вы получили с портала регистрации приложений.</span><span class="sxs-lookup"><span data-stu-id="92134-177">Replace `YOUR_APP_ID_HERE` with the **Application Id** you got from the App Registration Portal.</span></span>
1. <span data-ttu-id="92134-178">В интерфейсе командной строки перейдите в этот каталог и запустите следующую команду, чтобы установить требования.</span><span class="sxs-lookup"><span data-stu-id="92134-178">In your command-line interface (CLI), navigate to this directory and run the following command to install requirements.</span></span>

    ```Shell
    yarn install
    ```

## <a name="run-the-sample"></a><span data-ttu-id="92134-179">Запуск приложения</span><span class="sxs-lookup"><span data-stu-id="92134-179">Run the sample</span></span>

1. <span data-ttu-id="92134-180">Чтобы запустить приложение, запустите следующую команду в CLI.</span><span class="sxs-lookup"><span data-stu-id="92134-180">Run the following command in your CLI to start the application.</span></span>

    ```Shell
    yarn start
    ```

1. <span data-ttu-id="92134-181">В браузере перейдите в [Office.com](https://www.office.com/) и войдите.</span><span class="sxs-lookup"><span data-stu-id="92134-181">In your browser, go to [Office.com](https://www.office.com/) and sign in.</span></span> <span data-ttu-id="92134-182">Выберите **"Создать"** на левой панели инструментов, а затем выберите **"Таблица".**</span><span class="sxs-lookup"><span data-stu-id="92134-182">Select **Create** in the left-hand toolbar, then select **Spreadsheet**.</span></span>

    ![Снимок экрана: меню "Создать" Office.com](/tutorial/images/office-select-excel.png)

1. <span data-ttu-id="92134-184">Выберите **вкладку "Вставка"** и выберите **надстройки Office.**</span><span class="sxs-lookup"><span data-stu-id="92134-184">Select the **Insert** tab, then select **Office Add-ins**.</span></span>

1. <span data-ttu-id="92134-185">Выберите **"Отправить мою надстройка"** и выберите **"Обзор".**</span><span class="sxs-lookup"><span data-stu-id="92134-185">Select **Upload My Add-in**, then select **Browse**.</span></span> <span data-ttu-id="92134-186">Отправка **manifest.xml** файла.</span><span class="sxs-lookup"><span data-stu-id="92134-186">Upload your **manifest.xml** file.</span></span>

1. <span data-ttu-id="92134-187">Выберите **кнопку "Импорт календаря"** на **вкладке "Главная",** чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="92134-187">Select the **Import Calendar** button on the **Home** tab to open the taskpane.</span></span>

    ![Снимок экрана с кнопкой "Импорт календаря" на вкладке "Главная"](/tutorial/images/get-started.png)
