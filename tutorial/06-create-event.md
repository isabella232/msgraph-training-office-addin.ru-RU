---
ms.openlocfilehash: facdbb5c42e60e5bb0ee98b06ef68939a6010a67
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274319"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="da7cf-101">В этом разделе вы добавим возможность создания событий в календаре пользователя.</span><span class="sxs-lookup"><span data-stu-id="da7cf-101">In this section you will add the ability to create events on the user's calendar.</span></span>

## <a name="implement-the-api"></a><span data-ttu-id="da7cf-102">Реализация API</span><span class="sxs-lookup"><span data-stu-id="da7cf-102">Implement the API</span></span>

1. <span data-ttu-id="da7cf-103">Откройте **./src/api/graph.ts** и добавьте следующий код для реализации нового API события ( `POST /graph/newevent` ).</span><span class="sxs-lookup"><span data-stu-id="da7cf-103">Open **./src/api/graph.ts** and add the following code to implement a new event API (`POST /graph/newevent`).</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="CreateEventSnippet":::

1. <span data-ttu-id="da7cf-104">Откройте **./src/addin/taskpane.js** и добавьте следующую функцию для вызова нового API события.</span><span class="sxs-lookup"><span data-stu-id="da7cf-104">Open **./src/addin/taskpane.js** and add the following function to call the new event API.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="CreateEventSnippet":::

1. <span data-ttu-id="da7cf-105">Сохраните все изменения, перезапустите сервер и обновите области задач в Excel (закройте все открытые области задач и повторно откройте).</span><span class="sxs-lookup"><span data-stu-id="da7cf-105">Save all of your changes, restart the server, and refresh the task pane in Excel (close any open task panes and re-open).</span></span>

    ![Снимок экрана формы создания события](images/create-event-ui.png)

1. <span data-ttu-id="da7cf-107">Заполните форму и выберите **"Создать".**</span><span class="sxs-lookup"><span data-stu-id="da7cf-107">Fill in the form and choose **Create**.</span></span> <span data-ttu-id="da7cf-108">Убедитесь, что событие добавлено в календарь пользователя.</span><span class="sxs-lookup"><span data-stu-id="da7cf-108">Verify that the event is added to the user's calendar.</span></span>
