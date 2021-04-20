---
ms.openlocfilehash: bd5901525561f7e169f2b7c71a280883dac8c762
ms.sourcegitcommit: 8a65c826f6b229c287a782d784b6d9629aa5a3d0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/20/2021
ms.locfileid: "51899417"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="0dcdd-101">В этом упражнении вы будете включать Microsoft Graph в приложение.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-101">In this exercise you will incorporate Microsoft Graph into the application.</span></span> <span data-ttu-id="0dcdd-102">Для этого приложения вы будете использовать [библиотеку microsoft-graph-client](https://github.com/microsoftgraph/msgraph-sdk-javascript) для звонков в Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-102">For this application, you will use the [microsoft-graph-client](https://github.com/microsoftgraph/msgraph-sdk-javascript) library to make calls to Microsoft Graph.</span></span>

## <a name="get-calendar-events-from-outlook"></a><span data-ttu-id="0dcdd-103">Получение событий календаря из Outlook</span><span class="sxs-lookup"><span data-stu-id="0dcdd-103">Get calendar events from Outlook</span></span>

<span data-ttu-id="0dcdd-104">Начните с добавления API, чтобы получить представление [календаря](https://docs.microsoft.com/graph/api/user-list-calendarview) из календаря пользователя.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-104">Start by adding an API to get a [calendar view](https://docs.microsoft.com/graph/api/user-list-calendarview) from the user's calendar.</span></span>

1. <span data-ttu-id="0dcdd-105">Откройте **./src/api/graph.ts** и добавьте следующие утверждения в `import` верхнюю часть файла.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-105">Open **./src/api/graph.ts** and add the following `import` statements to the top of the file.</span></span>

    ```typescript
    import { zonedTimeToUtc } from 'date-fns-tz';
    import { findIana } from 'windows-iana';
    import * as graph from '@microsoft/microsoft-graph-client';
    import { Event, MailboxSettings } from 'microsoft-graph';
    import 'isomorphic-fetch';
    import { getTokenOnBehalfOf } from './auth';
    ```

1. <span data-ttu-id="0dcdd-106">Добавьте следующую функцию, чтобы инициализировать SDK Microsoft Graph и вернуть **клиента.**</span><span class="sxs-lookup"><span data-stu-id="0dcdd-106">Add the following function to initialize the Microsoft Graph SDK and return a **Client**.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetClientSnippet":::

1. <span data-ttu-id="0dcdd-107">Добавьте следующую функцию, чтобы получить часовой пояс пользователя из параметров почтовых ящиков и преобразовать это значение в идентификатор часовой зоны IANA.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-107">Add the following function to get the user's time zone from their mailbox settings, and to convert that value to an IANA time zone identifier.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetTimeZonesSnippet":::

1. <span data-ttu-id="0dcdd-108">Добавьте следующую функцию (ниже `const graphRouter = Router();` строки) для реализации конечной точки API `GET /graph/calendarview` ().</span><span class="sxs-lookup"><span data-stu-id="0dcdd-108">Add the following function (below the `const graphRouter = Router();` line) to implement an API endpoint (`GET /graph/calendarview`).</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetCalendarViewSnippet":::

    <span data-ttu-id="0dcdd-109">Рассмотрим, что делает этот код.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-109">Consider what this code does.</span></span>

    - <span data-ttu-id="0dcdd-110">Он получает часовой пояс пользователя и использует его, чтобы преобразовать начало и конец запрашиваемого представления календаря в значения UTC.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-110">It gets the user's time zone and uses that to convert the start and end of the requested calendar view into UTC values.</span></span>
    - <span data-ttu-id="0dcdd-111">Он делает `GET` конечную `/me/calendarview` точку API graph.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-111">It does a `GET` to the `/me/calendarview` Graph API endpoint.</span></span>
        - <span data-ttu-id="0dcdd-112">Она использует функцию, чтобы установить заготку, в результате чего время начала и окончания возвращаемого события должно быть скорректировано в `header` `Prefer: outlook.timezone` часовой пояс пользователя.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-112">It uses the `header` function to set the `Prefer: outlook.timezone` header, causing the start and end times of the returned events to be adjusted to the user's time zone.</span></span>
        - <span data-ttu-id="0dcdd-113">Она использует функцию для добавления параметров и параметров, устанавливая начало и `query` `startDateTime` конец представления `endDateTime` календаря.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-113">It uses the `query` function to add the `startDateTime` and `endDateTime` parameters, setting the start and end of the calendar view.</span></span>
        - <span data-ttu-id="0dcdd-114">Она использует `select` функцию для запроса только полей, используемых надстройки.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-114">It uses the `select` function to request only the fields used by the add-in.</span></span>
        - <span data-ttu-id="0dcdd-115">Она использует `orderby` функцию для сортировки результатов к началу работы.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-115">It uses the `orderby` function to sort the results by the start time.</span></span>
        - <span data-ttu-id="0dcdd-116">Он использует `top` функцию, чтобы ограничить результаты в одном запросе до 25.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-116">It uses the `top` function to limit the results in a single request to 25.</span></span>
    - <span data-ttu-id="0dcdd-117">Он использует **объект PageIteratorCallback** для итерации результатов и для дополнительных запросов, если доступны дополнительные страницы результатов. [](https://docs.microsoft.com/graph/sdks/paging)</span><span class="sxs-lookup"><span data-stu-id="0dcdd-117">It uses a **PageIteratorCallback** object to [iterate through the results](https://docs.microsoft.com/graph/sdks/paging) and to make additional requests if more pages of results are available.</span></span>

## <a name="update-the-ui"></a><span data-ttu-id="0dcdd-118">Обновление пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="0dcdd-118">Update the UI</span></span>

<span data-ttu-id="0dcdd-119">Теперь обновим области задач, чтобы пользователь указать дату начала и окончания представления календаря.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-119">Now let's update the task pane to allow the user to specify a start and end date for the calendar view.</span></span>

1. <span data-ttu-id="0dcdd-120">Откройте **./src/addin/taskpane.js** и замените существующую функцию `showMainUi` на следующую.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-120">Open **./src/addin/taskpane.js** and replace the existing `showMainUi` function with the following.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="MainUiSnippet":::

    <span data-ttu-id="0dcdd-121">Этот код добавляет простую форму, чтобы пользователь может указать дату начала и окончания.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-121">This code adds a simple form so the user can specify a start and end date.</span></span> <span data-ttu-id="0dcdd-122">Он также реализует вторую форму для создания нового события.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-122">It also implements a second form for creating a new event.</span></span> <span data-ttu-id="0dcdd-123">Эта форма пока ничего не выполняет, эту функцию можно реализовать в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-123">That form doesn't do anything for now, you'll implement that feature in the next section.</span></span>

1. <span data-ttu-id="0dcdd-124">Добавьте следующий код в файл, чтобы создать таблицу в активном таблице, содержащей события, извлеченные из представления календаря.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-124">Add the following code to the file to create a table in the active worksheet containing the events retrieved from the calendar view.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="WriteToSheetSnippet":::

1. <span data-ttu-id="0dcdd-125">Добавьте следующую функцию, чтобы вызвать API представления календаря.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-125">Add the following function to call the calendar view API.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="GetCalendarSnippet":::

1. <span data-ttu-id="0dcdd-126">Сохраните все изменения, перезапустите сервер и обновите области задач в Excel (закройте все открытые области задач и откройте заново).</span><span class="sxs-lookup"><span data-stu-id="0dcdd-126">Save all of your changes, restart the server, and refresh the task pane in Excel (close any open task panes and re-open).</span></span>

    ![Снимок экрана формы импорта](images/get-calendar-view-ui.png)

1. <span data-ttu-id="0dcdd-128">Выберите даты начала и окончания и выберите **Импорт**.</span><span class="sxs-lookup"><span data-stu-id="0dcdd-128">Choose start and end dates and choose **Import**.</span></span>

    ![Снимок экрана с таблицей событий](images/calendar-view-table.png)
