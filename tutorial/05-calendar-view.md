---
ms.openlocfilehash: ade142f1518d9bd56fa1472889721a4c8995bc3c
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274322"
---
<!-- markdownlint-disable MD002 MD041 -->

В этом упражнении вы включаете Microsoft Graph в приложение. Для этого приложения вы будете использовать клиентскую библиотеку [Microsoft-graph](https://github.com/microsoftgraph/msgraph-sdk-javascript) для вызова Microsoft Graph.

## <a name="get-calendar-events-from-outlook"></a>Получить события календаря из Outlook

Для начала добавьте API, чтобы получить представление [календаря](https://docs.microsoft.com/graph/api/user-list-calendarview) из календаря пользователя.

1. Откройте **файл ./src/api/graph.ts** и добавьте следующие утверждения в `import` верхнюю часть файла.

    ```typescript
    import { zonedTimeToUtc } from 'date-fns-tz';
    import { findOneIana } from 'windows-iana';
    import * as graph from '@microsoft/microsoft-graph-client';
    import { Event, MailboxSettings } from 'microsoft-graph';
    import 'isomorphic-fetch';
    import { getTokenOnBehalfOf } from './auth';
    ```

1. Добавьте следующую функцию для инициализации Microsoft Graph SDK и возврата **клиента.**

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetClientSnippet":::

1. Добавьте следующую функцию, чтобы получить часовой пояс пользователя из параметров почтового ящика и преобразовать это значение в идентификатор часовой пояс IANA.

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetTimeZonesSnippet":::

1. Добавьте следующую функцию (под `const graphRouter = Router();` строкой) для реализации конечной точки API ( `GET /graph/calendarview` ).

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetCalendarViewSnippet":::

    Подумайте, что делает этот код.

    - Он получает часовой пояс пользователя и использует его для преобразования начала и конца запрашиваемого представления календаря в значения UTC.
    - Он делает a `GET` для `/me/calendarview` конечной точки API Graph.
        - Она использует функцию для загона, в результате чего время начала и окончания возвращаемого события настраивается в часовом `header` `Prefer: outlook.timezone` поясе пользователя.
        - Функция используется для добавления параметров и параметров, устанавливая начало и конец `query` `startDateTime` представления `endDateTime` календаря.
        - Функция используется `select` для запроса только полей, используемых надстройой.
        - Функция используется `orderby` для сортировки результатов по времени начала.
        - Она использует `top` функцию, чтобы ограничить результаты в одном запросе до 25.
    - Он использует объект **PageIteratorCallback** для итерации результатов и для запроса дополнительных запросов, если доступно больше страниц результатов. [](https://docs.microsoft.com/graph/sdks/paging)

## <a name="update-the-ui"></a>Обновление пользовательского интерфейса

Теперь обновим области задач, чтобы пользователь указать дату начала и окончания для представления календаря.

1. Откройте **./src/addin/taskpane.js** замените существующую `showMainUi` функцию на следующую.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="MainUiSnippet":::

    В этом коде добавляется простая форма, в результате чего пользователь может указать дату начала и окончания. Он также реализует вторую форму для создания нового события. На данный момент эта форма ничего не выполняет, эта функция будет реализована в следующем разделе.

1. Добавьте в файл следующий код, чтобы создать таблицу на активном таблице, содержащей события, полученные из представления календаря.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="WriteToSheetSnippet":::

1. Добавьте следующую функцию для вызова API представления календаря.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="GetCalendarSnippet":::

1. Сохраните все изменения, перезапустите сервер и обновите области задач в Excel (закройте все открытые области задач и повторно откройте).

    ![Снимок экрана с формой импорта](images/get-calendar-view-ui.png)

1. Выберите даты начала и окончания и выберите **"Импорт".**

    ![Снимок экрана с таблицей событий](images/calendar-view-table.png)
