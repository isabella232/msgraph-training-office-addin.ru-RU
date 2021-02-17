---
ms.openlocfilehash: 2c323d61632c62c82af0561536656f1d68fe8938
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274376"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="ff69d-101">В этом руководстве рассказывается, как создать надстройки Office для Excel, использующие API Microsoft Graph для получения сведений календаря для пользователя.</span><span class="sxs-lookup"><span data-stu-id="ff69d-101">This tutorial teaches you how to build an Office Add-in for Excel that uses the Microsoft Graph API to retrieve calendar information for a user.</span></span>

> [!TIP]
> <span data-ttu-id="ff69d-102">Если вы предпочитаете просто скачать завершенный учебник, вы можете скачать или клонировать [репозиторий GitHub.](https://github.com/microsoftgraph/msgraph-training-office-addin)</span><span class="sxs-lookup"><span data-stu-id="ff69d-102">If you prefer to just download the completed tutorial, you can download or clone the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-office-addin).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ff69d-103">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="ff69d-103">Prerequisites</span></span>

<span data-ttu-id="ff69d-104">Перед началом этой демонстрации [](https://nodejs.org) необходимо установитьNode.js[yarn](https://yarnpkg.com/) на компьютере разработчика.</span><span class="sxs-lookup"><span data-stu-id="ff69d-104">Before you start this demo, you should have [Node.js](https://nodejs.org) and [Yarn](https://yarnpkg.com/) installed on your development machine.</span></span> <span data-ttu-id="ff69d-105">Если у вас нет Node.js или Yarn, посетите предыдущую ссылку для скачивания.</span><span class="sxs-lookup"><span data-stu-id="ff69d-105">If you do not have Node.js or Yarn, visit the previous link for download options.</span></span>

> [!NOTE]
> <span data-ttu-id="ff69d-106">Пользователям Windows может потребоваться установить Python и Visual Studio build Tools для поддержки модулей NPM, которые необходимо скомпилировать из C/C++.</span><span class="sxs-lookup"><span data-stu-id="ff69d-106">Windows users may need to install Python and Visual Studio Build Tools to support NPM modules that need to be compiled from C/C++.</span></span> <span data-ttu-id="ff69d-107">Установщик Node.js windows предоставляет возможность автоматической установки этих средств.</span><span class="sxs-lookup"><span data-stu-id="ff69d-107">The Node.js installer on Windows gives an option to automatically install these tools.</span></span> <span data-ttu-id="ff69d-108">Кроме того, вы можете следовать инструкциям по [https://github.com/nodejs/node-gyp#on-windows](https://github.com/nodejs/node-gyp#on-windows) этому пути.</span><span class="sxs-lookup"><span data-stu-id="ff69d-108">Alternatively, you can follow instructions at [https://github.com/nodejs/node-gyp#on-windows](https://github.com/nodejs/node-gyp#on-windows).</span></span>

<span data-ttu-id="ff69d-109">У вас также должна быть личная учетная запись Майкрософт с почтовым Outlook.com или учетная запись Майкрософт для работы или учебного заведения.</span><span class="sxs-lookup"><span data-stu-id="ff69d-109">You should also have either a personal Microsoft account with a mailbox on Outlook.com, or a Microsoft work or school account.</span></span> <span data-ttu-id="ff69d-110">Если у вас нет учетной записи Майкрософт, существует несколько вариантов получения бесплатной учетной записи:</span><span class="sxs-lookup"><span data-stu-id="ff69d-110">If you don't have a Microsoft account, there are a couple of options to get a free account:</span></span>

- <span data-ttu-id="ff69d-111">Вы можете [зарегистрироваться для новой личной учетной записи Майкрософт.](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1)</span><span class="sxs-lookup"><span data-stu-id="ff69d-111">You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).</span></span>
- <span data-ttu-id="ff69d-112">Вы можете зарегистрироваться в программе для разработчиков [Office 365,](https://developer.microsoft.com/office/dev-program) чтобы получить бесплатную подписку на Office 365.</span><span class="sxs-lookup"><span data-stu-id="ff69d-112">You can [sign up for the Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get a free Office 365 subscription.</span></span>

> [!NOTE]
> <span data-ttu-id="ff69d-113">Это руководство было написано с узлами версии 14.15.0 и Yarn версии 1.22.0.</span><span class="sxs-lookup"><span data-stu-id="ff69d-113">This tutorial was written with Node version 14.15.0 and Yarn version 1.22.0.</span></span> <span data-ttu-id="ff69d-114">Действия в этом руководстве могут работать с другими версиями, но не были протестированы.</span><span class="sxs-lookup"><span data-stu-id="ff69d-114">The steps in this guide may work with other versions, but that has not been tested.</span></span>

## <a name="feedback"></a><span data-ttu-id="ff69d-115">Отзывы</span><span class="sxs-lookup"><span data-stu-id="ff69d-115">Feedback</span></span>

<span data-ttu-id="ff69d-116">Поделитесь с этим учебником отзывами в [репозитории GitHub.](https://github.com/microsoftgraph/msgraph-training-office-addin)</span><span class="sxs-lookup"><span data-stu-id="ff69d-116">Please provide any feedback on this tutorial in the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-office-addin).</span></span>
