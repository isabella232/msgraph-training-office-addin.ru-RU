---
ms.openlocfilehash: a336095a238eefeae22dac86d29e3140e8d94bf1
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274337"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="13429-101">В этом упражнении вы создайте регистрацию нового веб-приложения Azure AD с помощью Центра администрирования Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="13429-101">In this exercise, you will create a new Azure AD web application registration using the Azure Active Directory admin center.</span></span>

1. <span data-ttu-id="13429-102">Откройте браузер и перейдите к [Центру администрирования Azure Active Directory](https://aad.portal.azure.com).</span><span class="sxs-lookup"><span data-stu-id="13429-102">Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com).</span></span> <span data-ttu-id="13429-103">Войдите с помощью **личной учетной записи** (т.е. учетной записи Microsoft) или **рабочей (учебной) учетной записи**.</span><span class="sxs-lookup"><span data-stu-id="13429-103">Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.</span></span>

1. <span data-ttu-id="13429-104">Выберите **Azure Active Directory** на панели навигации слева, затем выберите **Регистрация приложений** в разделе **Управление**.</span><span class="sxs-lookup"><span data-stu-id="13429-104">Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.</span></span>

    ![<span data-ttu-id="13429-105">Снимок экрана с регистрацией приложений</span><span class="sxs-lookup"><span data-stu-id="13429-105">A screenshot of the App registrations</span></span> ](images/app-registrations.png)

1. <span data-ttu-id="13429-106">Выберите **Новая регистрация**.</span><span class="sxs-lookup"><span data-stu-id="13429-106">Select **New registration**.</span></span> <span data-ttu-id="13429-107">На странице **Зарегистрировать приложение** задайте необходимые значения следующим образом.</span><span class="sxs-lookup"><span data-stu-id="13429-107">On the **Register an application** page, set the values as follows.</span></span>

    - <span data-ttu-id="13429-108">Введите **имя** `Office Add-in Graph Tutorial`.</span><span class="sxs-lookup"><span data-stu-id="13429-108">Set **Name** to `Office Add-in Graph Tutorial`.</span></span>
    - <span data-ttu-id="13429-109">Введите **поддерживаемые типы учетных записей** для **учетных записей в любом каталоге организаций и личных учетных записей Microsoft**.</span><span class="sxs-lookup"><span data-stu-id="13429-109">Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.</span></span>
    - <span data-ttu-id="13429-110">В разделе **URI адрес перенаправления** введите значение в первом раскрывающемся списке `Single-page application (SPA)` и задайте значение `https://localhost:3000/consent.html`.</span><span class="sxs-lookup"><span data-stu-id="13429-110">Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://localhost:3000/consent.html`.</span></span>

    ![Снимок экрана: страница "Регистрация приложения"](images/register-an-app.png)

1. <span data-ttu-id="13429-112">Нажмите **Зарегистрировать**.</span><span class="sxs-lookup"><span data-stu-id="13429-112">Select **Register**.</span></span> <span data-ttu-id="13429-113">На странице **учебника по Graph** для надстройки Office скопируйте значение ИД приложения **(клиента)** и сохраните его, оно потребуется вам на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="13429-113">On the **Office Add-in Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.</span></span>

    ![Снимок экрана: ИД нового приложения для регистрации](images/application-id.png)

1. <span data-ttu-id="13429-115">Выберите **Сертификаты и секреты** в разделе **Управление**.</span><span class="sxs-lookup"><span data-stu-id="13429-115">Select **Certificates & secrets** under **Manage**.</span></span> <span data-ttu-id="13429-116">Нажмите кнопку **Новый секрет клиента**.</span><span class="sxs-lookup"><span data-stu-id="13429-116">Select the **New client secret** button.</span></span> <span data-ttu-id="13429-117">Введите значение в **описании** и выберите один из параметров **"Срок действия истекает"** и выберите **"Добавить".**</span><span class="sxs-lookup"><span data-stu-id="13429-117">Enter a value in **Description** and select one of the options for **Expires** and select **Add**.</span></span>

1. <span data-ttu-id="13429-118">Скопируйте значение секрета клиента, а затем покиньте эту страницу.</span><span class="sxs-lookup"><span data-stu-id="13429-118">Copy the client secret value before you leave this page.</span></span> <span data-ttu-id="13429-119">Оно вам понадобится на следующем шаге.</span><span class="sxs-lookup"><span data-stu-id="13429-119">You will need it in the next step.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="13429-120">Это секрет клиента, он никогда не отображается еще раз, поэтому убедитесь, что вы скопировали его.</span><span class="sxs-lookup"><span data-stu-id="13429-120">This client secret is never shown again, so make sure you copy it now.</span></span>

1. <span data-ttu-id="13429-121">Выберите **разрешения API в** области **"Управление",** а затем **выберите "Добавить разрешение".**</span><span class="sxs-lookup"><span data-stu-id="13429-121">Select **API permissions** under **Manage**, then select **Add a permission**.</span></span>

1. <span data-ttu-id="13429-122">Выберите **Microsoft Graph,** а затем **делегирование разрешений.**</span><span class="sxs-lookup"><span data-stu-id="13429-122">Select **Microsoft Graph**, then **Delegated permissions**.</span></span>

1. <span data-ttu-id="13429-123">Выберите следующие разрешения, а затем выберите **"Добавить разрешения".**</span><span class="sxs-lookup"><span data-stu-id="13429-123">Select the following permissions, then select **Add permissions**.</span></span>

    - <span data-ttu-id="13429-124">**offline_access** — это позволит приложению обновлять маркеры доступа по истечении срока их действия.</span><span class="sxs-lookup"><span data-stu-id="13429-124">**offline_access** - this will allow the app to refresh access tokens when they expire.</span></span>
    - <span data-ttu-id="13429-125">**Calendars.ReadWrite** — позволяет приложению читать и записывать данные в календарь пользователя.</span><span class="sxs-lookup"><span data-stu-id="13429-125">**Calendars.ReadWrite** - this will allow the app to read and write to the user's calendar.</span></span>
    - <span data-ttu-id="13429-126">**MailboxSettings.Read** — позволяет приложению получить часовой пояс пользователя из параметров почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="13429-126">**MailboxSettings.Read** - this will allow the app to get the user's time zone from their mailbox settings.</span></span>

    ![Снимок экрана с настроенными разрешениями](images/configured-permissions.png)

## <a name="configure-office-add-in-single-sign-on"></a><span data-ttu-id="13429-128">Настройка единого входов для надстройки Office</span><span class="sxs-lookup"><span data-stu-id="13429-128">Configure Office Add-in single sign-on</span></span>

<span data-ttu-id="13429-129">В этом разделе вы обновим регистрацию приложения, чтобы поддерживать единый вход для [надстройки Office.](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)</span><span class="sxs-lookup"><span data-stu-id="13429-129">In this section you'll update the app registration to support [Office Add-in single sign-on (SSO)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).</span></span>

1. <span data-ttu-id="13429-130">Выберите **"Показать API"**.</span><span class="sxs-lookup"><span data-stu-id="13429-130">Select **Expose an API**.</span></span> <span data-ttu-id="13429-131">В разделе **"Области", определяемом этим API,** выберите **"Добавить область".**</span><span class="sxs-lookup"><span data-stu-id="13429-131">In the **Scopes defined by this API** section, select **Add a scope**.</span></span> <span data-ttu-id="13429-132">Когда вам будет предложено установить **URI** для ИД приложения, установите значение , заменив его `api://localhost:3000/YOUR_APP_ID_HERE` на `YOUR_APP_ID_HERE` ИД приложения.</span><span class="sxs-lookup"><span data-stu-id="13429-132">When prompted to set an **Application ID URI**, set the value to `api://localhost:3000/YOUR_APP_ID_HERE`, replacing `YOUR_APP_ID_HERE` with the application ID.</span></span> <span data-ttu-id="13429-133">Choose **Save and continue**.</span><span class="sxs-lookup"><span data-stu-id="13429-133">Choose **Save and continue**.</span></span>

1. <span data-ttu-id="13429-134">Заполните поля следующим образом и выберите **"Добавить область".**</span><span class="sxs-lookup"><span data-stu-id="13429-134">Fill in the fields as follows and select **Add scope**.</span></span>

    - <span data-ttu-id="13429-135">**Имя области:**`access_as_user`</span><span class="sxs-lookup"><span data-stu-id="13429-135">**Scope name:** `access_as_user`</span></span>
    - <span data-ttu-id="13429-136">**Кто может дать согласие?: администраторы и пользователи**</span><span class="sxs-lookup"><span data-stu-id="13429-136">**Who can consent?: Admins and users**</span></span>
    - <span data-ttu-id="13429-137">**Отображаемая имя согласия администратора:**`Access the app as the user`</span><span class="sxs-lookup"><span data-stu-id="13429-137">**Admin consent display name:** `Access the app as the user`</span></span>
    - <span data-ttu-id="13429-138">**Описание согласия администратора:**`Allows Office Add-ins to call the app's web APIs as the current user.`</span><span class="sxs-lookup"><span data-stu-id="13429-138">**Admin consent description:** `Allows Office Add-ins to call the app's web APIs as the current user.`</span></span>
    - <span data-ttu-id="13429-139">**Отображаемая имя согласия пользователя:**`Access the app as you`</span><span class="sxs-lookup"><span data-stu-id="13429-139">**User consent display name:** `Access the app as you`</span></span>
    - <span data-ttu-id="13429-140">**Описание согласия пользователя:**`Allows Office Add-ins to call the app's web APIs as you.`</span><span class="sxs-lookup"><span data-stu-id="13429-140">**User consent description:** `Allows Office Add-ins to call the app's web APIs as you.`</span></span>
    - <span data-ttu-id="13429-141">**Состояние: включено**</span><span class="sxs-lookup"><span data-stu-id="13429-141">**State: Enabled**</span></span>

    ![Снимок экрана с формой "Добавление области"](images/add-scope.png)

1. <span data-ttu-id="13429-143">В разделе **"Авторизованные клиентские приложения"** выберите **"Добавить клиентские приложения".**</span><span class="sxs-lookup"><span data-stu-id="13429-143">In the **Authorized client applications** section, select **Add a client application**.</span></span> <span data-ttu-id="13429-144">Введите ИД клиента из следующего списка, введите область в разрешенных разрешениях и выберите **"Добавить приложение".**</span><span class="sxs-lookup"><span data-stu-id="13429-144">Enter a client ID from the following list, enable the scope under **Authorized scopes**, and select **Add application**.</span></span> <span data-ttu-id="13429-145">Повторите этот процесс для каждого из клиентских ИД в списке.</span><span class="sxs-lookup"><span data-stu-id="13429-145">Repeat this process for each of the client IDs in the list.</span></span>

    - <span data-ttu-id="13429-146">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="13429-146">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="13429-147">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="13429-147">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="13429-148">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="13429-148">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="13429-149">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="13429-149">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span></span>
