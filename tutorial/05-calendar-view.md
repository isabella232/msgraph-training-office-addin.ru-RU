---
ms.openlocfilehash: ade142f1518d9bd56fa1472889721a4c8995bc3c
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274322"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="cf9be-101">В этом упражнении вы включаете Microsoft Graph в приложение.</span><span class="sxs-lookup"><span data-stu-id="cf9be-101">In this exercise you will incorporate Microsoft Graph into the application.</span></span> <span data-ttu-id="cf9be-102">Для этого приложения вы будете использовать клиентскую библиотеку [Microsoft-graph](https://github.com/microsoftgraph/msgraph-sdk-javascript) для вызова Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="cf9be-102">For this application, you will use the [microsoft-graph-client](https://github.com/microsoftgraph/msgraph-sdk-javascript) library to make calls to Microsoft Graph.</span></span>

## <a name="get-calendar-events-from-outlook"></a><span data-ttu-id="cf9be-103">Получить события календаря из Outlook</span><span class="sxs-lookup"><span data-stu-id="cf9be-103">Get calendar events from Outlook</span></span>

<span data-ttu-id="cf9be-104">Для начала добавьте API, чтобы получить представление [календаря](https://docs.microsoft.com/graph/api/user-list-calendarview) из календаря пользователя.</span><span class="sxs-lookup"><span data-stu-id="cf9be-104">Start by adding an API to get a [calendar view](https://docs.microsoft.com/graph/api/user-list-calendarview) from the user's calendar.</span></span>

1. <span data-ttu-id="cf9be-105">Откройте **файл ./src/api/graph.ts** и добавьте следующие утверждения в `import` верхнюю часть файла.</span><span class="sxs-lookup"><span data-stu-id="cf9be-105">Open **./src/api/graph.ts** and add the following `import` statements to the top of the file.</span></span>

    ```typescript
    import { zonedTimeToUtc } from 'date-fns-tz';
    import { findOneIana } from 'windows-iana';
    import * as graph from '@microsoft/microsoft-graph-client';
    import { Event, MailboxSettings } from 'microsoft-graph';
    import 'isomorphic-fetch';
    import { getTokenOnBehalfOf } from './auth';
    ```

1. <span data-ttu-id="cf9be-106">Добавьте следующую функцию для инициализации Microsoft Graph SDK и возврата **клиента.**</span><span class="sxs-lookup"><span data-stu-id="cf9be-106">Add the following function to initialize the Microsoft Graph SDK and return a **Client**.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetClientSnippet":::

1. <span data-ttu-id="cf9be-107">Добавьте следующую функцию, чтобы получить часовой пояс пользователя из параметров почтового ящика и преобразовать это значение в идентификатор часовой пояс IANA.</span><span class="sxs-lookup"><span data-stu-id="cf9be-107">Add the following function to get the user's time zone from their mailbox settings, and to convert that value to an IANA time zone identifier.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetTimeZonesSnippet":::

1. <span data-ttu-id="cf9be-108">Добавьте следующую функцию (под `const graphRouter = Router();` строкой) для реализации конечной точки API ( `GET /graph/calendarview` ).</span><span class="sxs-lookup"><span data-stu-id="cf9be-108">Add the following function (below the `const graphRouter = Router();` line) to implement an API endpoint (`GET /graph/calendarview`).</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetCalendarViewSnippet":::

    <span data-ttu-id="cf9be-109">Подумайте, что делает этот код.</span><span class="sxs-lookup"><span data-stu-id="cf9be-109">Consider what this code does.</span></span>

    - <span data-ttu-id="cf9be-110">Он получает часовой пояс пользователя и использует его для преобразования начала и конца запрашиваемого представления календаря в значения UTC.</span><span class="sxs-lookup"><span data-stu-id="cf9be-110">It gets the user's time zone and uses that to convert the start and end of the requested calendar view into UTC values.</span></span>
    - <span data-ttu-id="cf9be-111">Он делает a `GET` для `/me/calendarview` конечной точки API Graph.</span><span class="sxs-lookup"><span data-stu-id="cf9be-111">It does a `GET` to the `/me/calendarview` Graph API endpoint.</span></span>
        - <span data-ttu-id="cf9be-112">Она использует функцию для загона, в результате чего время начала и окончания возвращаемого события настраивается в часовом `header` `Prefer: outlook.timezone` поясе пользователя.</span><span class="sxs-lookup"><span data-stu-id="cf9be-112">It uses the `header` function to set the `Prefer: outlook.timezone` header, causing the start and end times of the returned events to be adjusted to the user's time zone.</span></span>
        - <span data-ttu-id="cf9be-113">Функция используется для добавления параметров и параметров, устанавливая начало и конец `query` `startDateTime` представления `endDateTime` календаря.</span><span class="sxs-lookup"><span data-stu-id="cf9be-113">It uses the `query` function to add the `startDateTime` and `endDateTime` parameters, setting the start and end of the calendar view.</span></span>
        - <span data-ttu-id="cf9be-114">Функция используется `select` для запроса только полей, используемых надстройой.</span><span class="sxs-lookup"><span data-stu-id="cf9be-114">It uses the `select` function to request only the fields used by the add-in.</span></span>
        - <span data-ttu-id="cf9be-115">Функция используется `orderby` для сортировки результатов по времени начала.</span><span class="sxs-lookup"><span data-stu-id="cf9be-115">It uses the `orderby` function to sort the results by the start time.</span></span>
        - <span data-ttu-id="cf9be-116">Она использует `top` функцию, чтобы ограничить результаты в одном запросе до 25.</span><span class="sxs-lookup"><span data-stu-id="cf9be-116">It uses the `top` function to limit the results in a single request to 25.</span></span>
    - <span data-ttu-id="cf9be-117">Он использует объект **PageIteratorCallback** для итерации результатов и для запроса дополнительных запросов, если доступно больше страниц результатов. [](https://docs.microsoft.com/graph/sdks/paging)</span><span class="sxs-lookup"><span data-stu-id="cf9be-117">It uses a **PageIteratorCallback** object to [iterate through the results](https://docs.microsoft.com/graph/sdks/paging) and to make additional requests if more pages of results are available.</span></span>

## <a name="update-the-ui"></a><span data-ttu-id="cf9be-118">Обновление пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="cf9be-118">Update the UI</span></span>

<span data-ttu-id="cf9be-119">Теперь обновим области задач, чтобы пользователь указать дату начала и окончания для представления календаря.</span><span class="sxs-lookup"><span data-stu-id="cf9be-119">Now let's update the task pane to allow the user to specify a start and end date for the calendar view.</span></span>

1. <span data-ttu-id="cf9be-120">Откройте **./src/addin/taskpane.js** замените существующую `showMainUi` функцию на следующую.</span><span class="sxs-lookup"><span data-stu-id="cf9be-120">Open **./src/addin/taskpane.js** and replace the existing `showMainUi` function with the following.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="MainUiSnippet":::

    <span data-ttu-id="cf9be-121">В этом коде добавляется простая форма, в результате чего пользователь может указать дату начала и окончания.</span><span class="sxs-lookup"><span data-stu-id="cf9be-121">This code adds a simple form so the user can specify a start and end date.</span></span> <span data-ttu-id="cf9be-122">Он также реализует вторую форму для создания нового события.</span><span class="sxs-lookup"><span data-stu-id="cf9be-122">It also implements a second form for creating a new event.</span></span> <span data-ttu-id="cf9be-123">На данный момент эта форма ничего не выполняет, эта функция будет реализована в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="cf9be-123">That form doesn't do anything for now, you'll implement that feature in the next section.</span></span>

1. <span data-ttu-id="cf9be-124">Добавьте в файл следующий код, чтобы создать таблицу на активном таблице, содержащей события, полученные из представления календаря.</span><span class="sxs-lookup"><span data-stu-id="cf9be-124">Add the following code to the file to create a table in the active worksheet containing the events retrieved from the calendar view.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="WriteToSheetSnippet":::

1. <span data-ttu-id="cf9be-125">Добавьте следующую функцию для вызова API представления календаря.</span><span class="sxs-lookup"><span data-stu-id="cf9be-125">Add the following function to call the calendar view API.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="GetCalendarSnippet":::

1. <span data-ttu-id="cf9be-126">Сохраните все изменения, перезапустите сервер и обновите области задач в Excel (закройте все открытые области задач и повторно откройте).</span><span class="sxs-lookup"><span data-stu-id="cf9be-126">Save all of your changes, restart the server, and refresh the task pane in Excel (close any open task panes and re-open).</span></span>

    ![Снимок экрана с формой импорта](images/get-calendar-view-ui.png)

1. <span data-ttu-id="cf9be-128">Выберите даты начала и окончания и выберите **"Импорт".**</span><span class="sxs-lookup"><span data-stu-id="cf9be-128">Choose start and end dates and choose **Import**.</span></span>

    ![Снимок экрана с таблицей событий](images/calendar-view-table.png)
