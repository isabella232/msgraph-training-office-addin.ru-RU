---
ms.openlocfilehash: bd5901525561f7e169f2b7c71a280883dac8c762
ms.sourcegitcommit: 8a65c826f6b229c287a782d784b6d9629aa5a3d0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/20/2021
ms.locfileid: "51899417"
---
<!-- markdownlint-disable MD002 MD041 -->

В этом упражнении вы будете включать Microsoft Graph в приложение. Для этого приложения вы будете использовать [библиотеку microsoft-graph-client](https://github.com/microsoftgraph/msgraph-sdk-javascript) для звонков в Microsoft Graph.

## <a name="get-calendar-events-from-outlook"></a>Получение событий календаря из Outlook

Начните с добавления API, чтобы получить представление [календаря](https://docs.microsoft.com/graph/api/user-list-calendarview) из календаря пользователя.

1. Откройте **./src/api/graph.ts** и добавьте следующие утверждения в `import` верхнюю часть файла.

    ```typescript
    import { zonedTimeToUtc } from 'date-fns-tz';
    import { findIana } from 'windows-iana';
    import * as graph from '@microsoft/microsoft-graph-client';
    import { Event, MailboxSettings } from 'microsoft-graph';
    import 'isomorphic-fetch';
    import { getTokenOnBehalfOf } from './auth';
    ```

1. Добавьте следующую функцию, чтобы инициализировать SDK Microsoft Graph и вернуть **клиента.**

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetClientSnippet":::

1. Добавьте следующую функцию, чтобы получить часовой пояс пользователя из параметров почтовых ящиков и преобразовать это значение в идентификатор часовой зоны IANA.

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetTimeZonesSnippet":::

1. Добавьте следующую функцию (ниже `const graphRouter = Router();` строки) для реализации конечной точки API `GET /graph/calendarview` ().

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetCalendarViewSnippet":::

    Рассмотрим, что делает этот код.

    - Он получает часовой пояс пользователя и использует его, чтобы преобразовать начало и конец запрашиваемого представления календаря в значения UTC.
    - Он делает `GET` конечную `/me/calendarview` точку API graph.
        - Она использует функцию, чтобы установить заготку, в результате чего время начала и окончания возвращаемого события должно быть скорректировано в `header` `Prefer: outlook.timezone` часовой пояс пользователя.
        - Она использует функцию для добавления параметров и параметров, устанавливая начало и `query` `startDateTime` конец представления `endDateTime` календаря.
        - Она использует `select` функцию для запроса только полей, используемых надстройки.
        - Она использует `orderby` функцию для сортировки результатов к началу работы.
        - Он использует `top` функцию, чтобы ограничить результаты в одном запросе до 25.
    - Он использует **объект PageIteratorCallback** для итерации результатов и для дополнительных запросов, если доступны дополнительные страницы результатов. [](https://docs.microsoft.com/graph/sdks/paging)

## <a name="update-the-ui"></a>Обновление пользовательского интерфейса

Теперь обновим области задач, чтобы пользователь указать дату начала и окончания представления календаря.

1. Откройте **./src/addin/taskpane.js** и замените существующую функцию `showMainUi` на следующую.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="MainUiSnippet":::

    Этот код добавляет простую форму, чтобы пользователь может указать дату начала и окончания. Он также реализует вторую форму для создания нового события. Эта форма пока ничего не выполняет, эту функцию можно реализовать в следующем разделе.

1. Добавьте следующий код в файл, чтобы создать таблицу в активном таблице, содержащей события, извлеченные из представления календаря.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="WriteToSheetSnippet":::

1. Добавьте следующую функцию, чтобы вызвать API представления календаря.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="GetCalendarSnippet":::

1. Сохраните все изменения, перезапустите сервер и обновите области задач в Excel (закройте все открытые области задач и откройте заново).

    ![Снимок экрана формы импорта](images/get-calendar-view-ui.png)

1. Выберите даты начала и окончания и выберите **Импорт**.

    ![Снимок экрана с таблицей событий](images/calendar-view-table.png)
