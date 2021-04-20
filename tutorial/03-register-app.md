---
ms.openlocfilehash: 89bc4baff47ae895c7c0dfd34a8f8d4d0da091f5
ms.sourcegitcommit: 8a65c826f6b229c287a782d784b6d9629aa5a3d0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/20/2021
ms.locfileid: "51899439"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="c99ec-101">В этом упражнении будет создаваться новая регистрация веб-приложений Azure AD с помощью центра администрирования Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="c99ec-101">In this exercise, you will create a new Azure AD web application registration using the Azure Active Directory admin center.</span></span>

1. <span data-ttu-id="c99ec-102">Откройте браузер и перейдите в [Центр администрирования Azure Active Directory](https://aad.portal.azure.com).</span><span class="sxs-lookup"><span data-stu-id="c99ec-102">Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com).</span></span> <span data-ttu-id="c99ec-103">Войдите с помощью **личной учетной записи** (т.е. учетной записи Microsoft) или **рабочей (учебной) учетной записи**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-103">Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.</span></span>

1. <span data-ttu-id="c99ec-104">Выберите **Azure Active Directory** на панели навигации слева, затем выберите **Регистрация приложений** в разделе **Управление**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-104">Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.</span></span>

    ![<span data-ttu-id="c99ec-105">Снимок экрана регистрации приложения</span><span class="sxs-lookup"><span data-stu-id="c99ec-105">A screenshot of the App registrations</span></span> ](images/app-registrations.png)

1. <span data-ttu-id="c99ec-106">Выберите **Новая регистрация**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-106">Select **New registration**.</span></span> <span data-ttu-id="c99ec-107">На странице **Зарегистрировать приложение** задайте необходимые значения следующим образом.</span><span class="sxs-lookup"><span data-stu-id="c99ec-107">On the **Register an application** page, set the values as follows.</span></span>

    - <span data-ttu-id="c99ec-108">Введите **имя** `Office Add-in Graph Tutorial`.</span><span class="sxs-lookup"><span data-stu-id="c99ec-108">Set **Name** to `Office Add-in Graph Tutorial`.</span></span>
    - <span data-ttu-id="c99ec-109">Введите **поддерживаемые типы учетных записей** для **учетных записей в любом каталоге организаций и личных учетных записей Microsoft**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-109">Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.</span></span>
    - <span data-ttu-id="c99ec-110">В разделе **URI адрес перенаправления** введите значение в первом раскрывающемся списке `Single-page application (SPA)` и задайте значение `https://localhost:3000/consent.html`.</span><span class="sxs-lookup"><span data-stu-id="c99ec-110">Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://localhost:3000/consent.html`.</span></span>

    ![Снимок экрана со страницей регистрации приложения](images/register-an-app.png)

1. <span data-ttu-id="c99ec-112">Нажмите **Зарегистрировать**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-112">Select **Register**.</span></span> <span data-ttu-id="c99ec-113">На странице **Учебник по** надстройке Office скопируйте значение ID приложения **(клиента)** и сохраните его, оно потребуется на следующем шаге.</span><span class="sxs-lookup"><span data-stu-id="c99ec-113">On the **Office Add-in Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.</span></span>

    ![Снимок экрана с идентификатором приложения для новой регистрации](images/application-id.png)

1. <span data-ttu-id="c99ec-115">Выберите пункт **Проверка подлинности** в разделе **Управление**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-115">Select **Authentication** under **Manage**.</span></span> <span data-ttu-id="c99ec-116">Найдите **раздел Неявный грант** и впускаем **маркеры Access** и **маркеры ID.**</span><span class="sxs-lookup"><span data-stu-id="c99ec-116">Locate the **Implicit grant** section and enable **Access tokens** and **ID tokens**.</span></span> <span data-ttu-id="c99ec-117">Нажмите **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-117">Select **Save**.</span></span>

    ![Снимок экрана: раздел "Неявное предоставление разрешения"](./images/aad-implicit-grant.png)

1. <span data-ttu-id="c99ec-119">Выберите **Сертификаты и секреты** в разделе **Управление**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-119">Select **Certificates & secrets** under **Manage**.</span></span> <span data-ttu-id="c99ec-120">Нажмите кнопку **Новый секрет клиента**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-120">Select the **New client secret** button.</span></span> <span data-ttu-id="c99ec-121">Введите значение в поле **Описание**, выберите один из параметров **Срок действия** и нажмите **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-121">Enter a value in **Description** and select one of the options for **Expires** and select **Add**.</span></span>

1. <span data-ttu-id="c99ec-122">Скопируйте значение секрета клиента, а затем покиньте эту страницу.</span><span class="sxs-lookup"><span data-stu-id="c99ec-122">Copy the client secret value before you leave this page.</span></span> <span data-ttu-id="c99ec-123">Оно вам понадобится на следующем шаге.</span><span class="sxs-lookup"><span data-stu-id="c99ec-123">You will need it in the next step.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="c99ec-124">Это секрет клиента, он никогда не отображается еще раз, поэтому убедитесь, что вы скопировали его.</span><span class="sxs-lookup"><span data-stu-id="c99ec-124">This client secret is never shown again, so make sure you copy it now.</span></span>

1. <span data-ttu-id="c99ec-125">Выберите **разрешения API в** статье **Управление,** а затем **добавьте разрешение.**</span><span class="sxs-lookup"><span data-stu-id="c99ec-125">Select **API permissions** under **Manage**, then select **Add a permission**.</span></span>

1. <span data-ttu-id="c99ec-126">Выберите **Microsoft Graph,** затем **делегирование разрешений.**</span><span class="sxs-lookup"><span data-stu-id="c99ec-126">Select **Microsoft Graph**, then **Delegated permissions**.</span></span>

1. <span data-ttu-id="c99ec-127">Выберите следующие разрешения, а затем **добавьте разрешения.**</span><span class="sxs-lookup"><span data-stu-id="c99ec-127">Select the following permissions, then select **Add permissions**.</span></span>

    - <span data-ttu-id="c99ec-128">**offline_access** — это позволит приложению обновить маркеры доступа по истечении срока действия.</span><span class="sxs-lookup"><span data-stu-id="c99ec-128">**offline_access** - this will allow the app to refresh access tokens when they expire.</span></span>
    - <span data-ttu-id="c99ec-129">**Calendars.ReadWrite** — это позволит приложению читать и писать в календарь пользователя.</span><span class="sxs-lookup"><span data-stu-id="c99ec-129">**Calendars.ReadWrite** - this will allow the app to read and write to the user's calendar.</span></span>
    - <span data-ttu-id="c99ec-130">**MailboxSettings.Read** — это позволит приложению получать часовой пояс пользователя из параметров почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="c99ec-130">**MailboxSettings.Read** - this will allow the app to get the user's time zone from their mailbox settings.</span></span>

    ![Снимок экрана настроенных разрешений](images/configured-permissions.png)

## <a name="configure-office-add-in-single-sign-on"></a><span data-ttu-id="c99ec-132">Настройка единого входного знака Надстройки Office</span><span class="sxs-lookup"><span data-stu-id="c99ec-132">Configure Office Add-in single sign-on</span></span>

<span data-ttu-id="c99ec-133">В этом разделе вы обновим регистрацию приложения для поддержки единого входного знака [Office Add-in (SSO).](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)</span><span class="sxs-lookup"><span data-stu-id="c99ec-133">In this section you'll update the app registration to support [Office Add-in single sign-on (SSO)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).</span></span>

1. <span data-ttu-id="c99ec-134">Выберите **Expose aPI**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-134">Select **Expose an API**.</span></span> <span data-ttu-id="c99ec-135">В **области, определенные этим разделом API,** выберите **Добавить область**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-135">In the **Scopes defined by this API** section, select **Add a scope**.</span></span> <span data-ttu-id="c99ec-136">При запросе на набор **URI ID** приложения задай значение `api://localhost:3000/YOUR_APP_ID_HERE` , заменив его `YOUR_APP_ID_HERE` ИД приложения.</span><span class="sxs-lookup"><span data-stu-id="c99ec-136">When prompted to set an **Application ID URI**, set the value to `api://localhost:3000/YOUR_APP_ID_HERE`, replacing `YOUR_APP_ID_HERE` with the application ID.</span></span> <span data-ttu-id="c99ec-137">Выберите **Сохранить и продолжить**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-137">Choose **Save and continue**.</span></span>

1. <span data-ttu-id="c99ec-138">Заполните поля следующим образом и выберите **область Добавить**.</span><span class="sxs-lookup"><span data-stu-id="c99ec-138">Fill in the fields as follows and select **Add scope**.</span></span>

    - <span data-ttu-id="c99ec-139">**Имя области:**`access_as_user`</span><span class="sxs-lookup"><span data-stu-id="c99ec-139">**Scope name:** `access_as_user`</span></span>
    - <span data-ttu-id="c99ec-140">**Кто может дать согласие?: Администраторы и пользователи**</span><span class="sxs-lookup"><span data-stu-id="c99ec-140">**Who can consent?: Admins and users**</span></span>
    - <span data-ttu-id="c99ec-141">**Имя отображения согласия администратора:**`Access the app as the user`</span><span class="sxs-lookup"><span data-stu-id="c99ec-141">**Admin consent display name:** `Access the app as the user`</span></span>
    - <span data-ttu-id="c99ec-142">**Описание согласия администратора:**`Allows Office Add-ins to call the app's web APIs as the current user.`</span><span class="sxs-lookup"><span data-stu-id="c99ec-142">**Admin consent description:** `Allows Office Add-ins to call the app's web APIs as the current user.`</span></span>
    - <span data-ttu-id="c99ec-143">**Имя отображения согласия пользователя:**`Access the app as you`</span><span class="sxs-lookup"><span data-stu-id="c99ec-143">**User consent display name:** `Access the app as you`</span></span>
    - <span data-ttu-id="c99ec-144">**Описание согласия пользователя:**`Allows Office Add-ins to call the app's web APIs as you.`</span><span class="sxs-lookup"><span data-stu-id="c99ec-144">**User consent description:** `Allows Office Add-ins to call the app's web APIs as you.`</span></span>
    - <span data-ttu-id="c99ec-145">**Состояние: Включено**</span><span class="sxs-lookup"><span data-stu-id="c99ec-145">**State: Enabled**</span></span>

    ![Снимок экрана формы Добавить область](images/add-scope.png)

1. <span data-ttu-id="c99ec-147">В разделе **Авторизованные клиентские приложения** выберите **Добавление клиентского приложения.**</span><span class="sxs-lookup"><span data-stu-id="c99ec-147">In the **Authorized client applications** section, select **Add a client application**.</span></span> <span data-ttu-id="c99ec-148">Введите клиентский ИД из следующего списка, введите область в уполномоченных сферах **и** выберите **приложение Добавить.**</span><span class="sxs-lookup"><span data-stu-id="c99ec-148">Enter a client ID from the following list, enable the scope under **Authorized scopes**, and select **Add application**.</span></span> <span data-ttu-id="c99ec-149">Повторите этот процесс для каждого из клиентских ИД в списке.</span><span class="sxs-lookup"><span data-stu-id="c99ec-149">Repeat this process for each of the client IDs in the list.</span></span>

    - <span data-ttu-id="c99ec-150">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="c99ec-150">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="c99ec-151">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="c99ec-151">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="c99ec-152">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="c99ec-152">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="c99ec-153">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="c99ec-153">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span></span>
