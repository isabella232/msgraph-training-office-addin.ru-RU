---
ms.openlocfilehash: 2c323d61632c62c82af0561536656f1d68fe8938
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274376"
---
<!-- markdownlint-disable MD002 MD041 -->

В этом руководстве рассказывается, как создать надстройки Office для Excel, использующие API Microsoft Graph для получения сведений календаря для пользователя.

> [!TIP]
> Если вы предпочитаете просто скачать завершенный учебник, вы можете скачать или клонировать [репозиторий GitHub.](https://github.com/microsoftgraph/msgraph-training-office-addin)

## <a name="prerequisites"></a>Необходимые компоненты

Перед началом этой демонстрации [](https://nodejs.org) необходимо установитьNode.js[yarn](https://yarnpkg.com/) на компьютере разработчика. Если у вас нет Node.js или Yarn, посетите предыдущую ссылку для скачивания.

> [!NOTE]
> Пользователям Windows может потребоваться установить Python и Visual Studio build Tools для поддержки модулей NPM, которые необходимо скомпилировать из C/C++. Установщик Node.js windows предоставляет возможность автоматической установки этих средств. Кроме того, вы можете следовать инструкциям по [https://github.com/nodejs/node-gyp#on-windows](https://github.com/nodejs/node-gyp#on-windows) этому пути.

У вас также должна быть личная учетная запись Майкрософт с почтовым Outlook.com или учетная запись Майкрософт для работы или учебного заведения. Если у вас нет учетной записи Майкрософт, существует несколько вариантов получения бесплатной учетной записи:

- Вы можете [зарегистрироваться для новой личной учетной записи Майкрософт.](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1)
- Вы можете зарегистрироваться в программе для разработчиков [Office 365,](https://developer.microsoft.com/office/dev-program) чтобы получить бесплатную подписку на Office 365.

> [!NOTE]
> Это руководство было написано с узлами версии 14.15.0 и Yarn версии 1.22.0. Действия в этом руководстве могут работать с другими версиями, но не были протестированы.

## <a name="feedback"></a>Отзывы

Поделитесь с этим учебником отзывами в [репозитории GitHub.](https://github.com/microsoftgraph/msgraph-training-office-addin)
