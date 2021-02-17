---
ms.openlocfilehash: facdbb5c42e60e5bb0ee98b06ef68939a6010a67
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274319"
---
<!-- markdownlint-disable MD002 MD041 -->

В этом разделе вы добавим возможность создания событий в календаре пользователя.

## <a name="implement-the-api"></a>Реализация API

1. Откройте **./src/api/graph.ts** и добавьте следующий код для реализации нового API события ( `POST /graph/newevent` ).

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="CreateEventSnippet":::

1. Откройте **./src/addin/taskpane.js** и добавьте следующую функцию для вызова нового API события.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="CreateEventSnippet":::

1. Сохраните все изменения, перезапустите сервер и обновите области задач в Excel (закройте все открытые области задач и повторно откройте).

    ![Снимок экрана формы создания события](images/create-event-ui.png)

1. Заполните форму и выберите **"Создать".** Убедитесь, что событие добавлено в календарь пользователя.
