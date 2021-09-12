/* 1.0.0 отправляет переменные среды

cscript env.send.min.js <mode> <container> [<output>...] \\ [<input>...]

<mode>      - Режим отправки переменных (заглавное написание выполняет только эмуляцию).
    link    - Отправляет в обычный ярлык.
    ldap    - Отправляет в объект active directory.
<container> - Путь к папке или guid (допускается пустое значение).
<output>    - Изменяемые свойства объекта в формате key=value c подстановкой переменных %ENV%.
              Первое свойство считается оснавным и по его значению осуществляется поиск объектов.
              Для режима link обязательно наличие свойств name и targetpath, а в свойстве arguments
              одинарные кавычки заменяются на двойные.
<input>     - Значения по умолчанию для переменных среды в формате key=value.

*/

var send = new App({
    argWrap: '"',                                       // основное обрамление аргументов
    altWrap: "'",                                       // альтернативное обрамление аргументов
    argDelim: " ",                                      // разделитель значений агрументов
    lineDelim: "\r\n",                                  // разделитель строк со значениями
    extDelim: ".",                                      // разделитель расширений файлов
    keyDelim: "=",                                      // разделитель ключа от значения
    putDelim: "\\\\",                                   // разделитель потоков параметров
    linkExt: "lnk",                                     // расширение для ярлычков
    envType: "Process"                                  // тип изменяемого переменного окружения
});

// подключаем зависимые свойства приложения
(function (wsh, app, undefined) {
    app.lib.extend(app, {
        fun: {// зависимые функции частного назначения
        },
        init: function () {// функция инициализации приложения
            var key, value, list, mode, container, index, length, primary, file, files, path,
                items, item, isDelim, isMatch, isChange, isIgnoreCase, isEmulate, shell, fso,
                query, id, line, lines, output = {}, input = {}, data = {}, error = 0;

            shell = new ActiveXObject("WScript.Shell");
            fso = new ActiveXObject("Scripting.FileSystemObject");
            // получаем основные параметры
            if (!error) {// если нет ошибок
                length = wsh.arguments.length;// получаем длину
                for (index = 0; index < Math.min(length, 2); index++) {
                    value = wsh.arguments.item(index);// получаем значение
                    switch (index) {// поддерживаемые параметры
                        case 0:// режим работы
                            isEmulate = value == value.toUpperCase();
                            mode = value.toLowerCase();
                            break;
                        case 1:// контейнер
                            container = value;
                            break;
                    };
                };
            };
            // получаем исходящие параметры
            if (!error) {// если нет ошибок
                isDelim = false;// сбрасываем значение
                while (index < length && !isDelim && !error) {
                    value = wsh.arguments.item(index);// получаем значение
                    if (app.val.putDelim != value) {// если не разделитель потоков
                        key = app.lib.strim(value, null, app.val.keyDelim, false, false).toLowerCase();
                        if (key) {// если параметр имеет нужный формат
                            value = app.lib.strim(value, app.val.keyDelim, null, false, false);
                            list = value.split(app.val.argWrap);// вспомогательная переменная
                            if (3 == list.length && !list[0] && !list[2]) value = list[1];
                            if (!primary) primary = key;
                            output[key] = value;
                        } else error = 1;
                    } else isDelim = true;
                    index++;
                };
            };
            // получаем входящие параметры
            if (!error) {// если нет ошибок
                while (index < length && !error) {
                    value = wsh.arguments.item(index);// получаем значение
                    key = app.lib.strim(value, null, app.val.keyDelim, false, false);
                    if (key) {// если параметр имеет нужный формат
                        value = app.lib.strim(value, app.val.keyDelim, null, false, false);
                        list = value.split(app.val.argWrap);// вспомогательная переменная
                        if (3 == list.length && !list[0] && !list[2]) value = list[1];
                        input[key] = value;
                    } else error = 2;
                    index++;
                };
            };
            // работаем с переменными среды
            if (!error) {// если нет ошибок
                items = shell.environment(app.val.envType);
                // добавляем входящие параметры в переменные среды
                for (var key in input) {// пробегаемся по входящим параметрам
                    value = input[key];// получаем очередное значение
                    if (!items(key)) {// если значение ещё не задано
                        setEnv(items, key, input[key]);
                    };
                };
                // подставляем переменные среды в исходящие параметры
                for (var key in output) {// пробегаемся по входящим параметрам
                    value = output[key];// получаем очередное значение
                    value = shell.expandEnvironmentStrings(value);
                    output[key] = value;
                };
            };
            // работаем в зависимости от режима
            switch (mode) {// поддерживаемые режимы
                case "link":// ссылка
                    // проверяем обязательные параметры
                    if (!error) {// если нет ошибок
                        if (// множественное условие
                            output.name && output.targetpath && primary
                        ) {// если проверка пройдена
                            // корректируем обёртку аргументов
                            if (output.arguments) {// если заданы аргументы
                                list = output.arguments.split(app.val.altWrap);
                                output.arguments = list.join(app.val.argWrap);
                            };
                        } else error = 4;
                    };
                    // получаем контейнер
                    if (!error) {// если нет ошибок
                        path = fso.getAbsolutePathName(container);
                        if (fso.folderExists(path)) {// если контейнер существует
                            container = fso.getFolder(path);
                        } else error = 5;
                    };
                    // выполняем поиск целевых объектов
                    if (!error) {// если нет ошибок
                        items = [];// массив целевых объектов
                        files = new Enumerator(container.files);
                        while (!files.atEnd()) {// пока не достигнут конец
                            file = files.item();// получаем очередной элимент коллекции
                            files.moveNext();// переходим к следующему элименту
                            if (!app.lib.compare(app.val.linkExt, fso.getExtensionName(file.path), true)) {
                                item = shell.createShortcut(file.path);// получаем ярлычёк
                                switch (primary) {// проверка свойств объекта
                                    case "name":// имя файла
                                        isMatch = !app.lib.compare(
                                            output[primary].split(app.val.argDelim)[0],
                                            fso.getBaseName(file.path).split(app.val.argDelim)[0],
                                            true
                                        );
                                        break;
                                    default:// стандартное свойство
                                        isMatch = !app.lib.compare(
                                            output[primary],
                                            item[primary],
                                            true
                                        );
                                };
                                if (isMatch) items.push(item);
                            };
                        };
                    };
                    // создаём целевой объект
                    if (!error && !items.length) {// если нужно выполнить
                        value = output.name + app.val.extDelim + app.val.linkExt;
                        path = fso.buildPath(container.path, value);
                        if (!isEmulate) {// если не эмуляция
                            try {// пробуем выполнить комманду
                                item = shell.createShortcut(path);
                                items.push(item);
                            } catch (e) {// если ошибка
                                error = 6;
                            };
                        } else {// если используется эмуляция
                            item = { fullName: path };
                            id = fso.getBaseName(item.fullName);
                            if (!(id in data)) data[id] = { action: "create", change: {}, show: true };
                            items.push(item);
                        };
                    };
                    // изменяем целевые объекты
                    if (!error) {// если нет ошибок
                        length = items.length;// получаем длину
                        for (index = 0; index < Math.min(length, 1) && !error; index++) {
                            item = items[index];// получаем очередной элимент
                            isChange = false;// сбрасываем значение
                            id = fso.getBaseName(item.fullName);
                            if (!(id in data)) data[id] = { action: "edit", change: {}, show: false };
                            // изменяем стандартные свойства
                            for (var key in output) {// пробигаемся по свойствам
                                isIgnoreCase = false;// по умолчанию
                                switch (key) {// изменение свойств объекта
                                    case "name":// имя файла
                                        break;
                                    case "targetpath":// целевой объект
                                        isIgnoreCase = true;
                                    default:// стандартное свойство
                                        if (!error) {// если нет ошибок
                                            isMatch = !app.lib.compare(
                                                output[key],
                                                item[key],
                                                isIgnoreCase
                                            );
                                            if (!isMatch) {// если не совпадает
                                                if (!isEmulate) {// если не эмуляция
                                                    try {// пробуем выполнить комманду
                                                        item[key] = output[key];
                                                        isChange = true;
                                                    } catch (e) {// если ошибка
                                                        error = 7;
                                                    };
                                                } else {// если используется эмуляция
                                                    data[id].change[key] = output[key];
                                                    data[id].show = true;
                                                };
                                            };
                                        };
                                };
                            };
                            // сохраняем изменения
                            if (!error && isChange) {// если нужно выполнить
                                if (!isEmulate) {// если не эмуляция
                                    try {// пробуем выполнить комманду
                                        item.save();
                                    } catch (e) {// если ошибка
                                        error = 8;
                                    };
                                };
                            };
                            // изменяем не стандартные свойства
                            for (var key in output) {// пробигаемся по свойствам
                                isIgnoreCase = false;// по умолчанию
                                switch (key) {// изменение свойств объекта
                                    case "name":// имя файла
                                        if (!error) {// если нет ошибок
                                            isMatch = !app.lib.compare(
                                                fso.getBaseName(item.fullName),
                                                output[key],
                                                isIgnoreCase
                                            );
                                            if (!isMatch) {// если не совпадает
                                                value = output[key] + app.val.extDelim + app.val.linkExt;
                                                path = fso.buildPath(container.path, value);
                                                if (!isEmulate) {// если не эмуляция
                                                    try {// пробуем выполнить комманду
                                                        if (fso.fileExists(path)) fso.deleteFile(path, true);
                                                        fso.moveFile(item.fullName, path);
                                                    } catch (e) {// если ошибка
                                                        error = 9;
                                                    };
                                                } else {// если используется эмуляция
                                                    data[id].change[key] = output[key];
                                                    data[id].show = true;
                                                };
                                            };
                                        };
                                        break;
                                };
                            };
                        };
                    };
                    // удаляем лишнии целевые объекты
                    if (!error) {// если нет ошибок
                        value = output.name + app.val.extDelim + app.val.linkExt;
                        path = fso.buildPath(container.path, value);
                        while (index < length && !error) {
                            item = items[index];// получаем очередной элимент
                            id = fso.getBaseName(item.fullName);
                            isMatch = !app.lib.compare(
                                fso.getBaseName(item.fullName),
                                output.name,
                                true
                            );
                            if (!isEmulate) {// если не эмуляция
                                if (!isMatch) {// если не совпадает
                                    try {// пробуем выполнить комманду
                                        fso.deleteFile(item.fullName, true);
                                    } catch (e) {// если ошибка
                                        error = 10;
                                    };
                                };
                            } else {// если используется эмуляция
                                if (!(id in data)) data[id] = { action: "delete", change: {}, show: true };
                            };
                            index++;
                        };
                    };
                    break;
                case "ldap":// домен
                    // проверяем обязательные параметры
                    if (!error) {// если нет ошибок
                        if (// множественное условие
                            (!container || app.lib.validate(container, "guid")) && primary
                        ) {// если проверка пройдена
                        } else error = 4;
                    };
                    // получаем контейнер
                    if (!error) {// если нет ошибок
                        container = app.wsh.getLDAP(container)[0];
                        if (container) {// если контейнер существует
                        } else error = 5;
                    };
                    // выполняем поиск целевых объектов
                    if (!error) {// если нет ошибок
                        query = "WHERE " + primary + " = '" + output[primary] + "'";
                        items = app.wsh.getLDAP(query, container);
                    };
                    // изменяем целевые объекты
                    if (!error) {// если нет ошибок
                        length = items.length;// получаем длину
                        for (index = 0; index < length && !error; index++) {
                            item = items[index];// получаем очередной элимент
                            isChange = false;// сбрасываем значение
                            id = item.cn;// идентификатор для эмуляции
                            if (!(id in data)) data[id] = { action: "edit", change: {}, show: false };
                            // изменяем стандартные свойства
                            for (var key in output) {// пробигаемся по свойствам
                                isIgnoreCase = false;// по умолчанию
                                switch (key) {// изменение свойств объекта
                                    case primary:// ключевое свойство
                                        break;
                                    default:// стандартное свойство
                                        if (!error) {// если нет ошибок
                                            isMatch = !app.lib.compare(
                                                output[key],
                                                item[key],
                                                isIgnoreCase
                                            );
                                            if (!isMatch) {// если не совпадает
                                                if (!isEmulate) {// если не эмуляция
                                                    try {// пробуем выполнить комманду
                                                        value = output[key];
                                                        item.put(key, value);
                                                        isChange = true;
                                                    } catch (e) {// если ошибка
                                                        error = 7;
                                                    };
                                                } else {// если используется эмуляция
                                                    data[id].change[key] = output[key];
                                                    data[id].show = true;
                                                };
                                            };
                                        };
                                };
                            };
                            // сохраняем изменения
                            if (!error && isChange) {// если нужно выполнить
                                if (!isEmulate) {// если не эмуляция
                                    try {// пробуем выполнить комманду
                                        item.setInfo();
                                    } catch (e) {// если ошибка
                                        error = 8;
                                    };
                                };
                            };
                        };
                    };
                    break;
                default:// не поддерживаемый режим
                    if (!error) error = 3;
            };
            // выводим информацию по эмуляции
            if (!error && isEmulate) {// если нужно выполнить
                lines = [];// сбрасываем значение
                // формируем строки для вывода
                for (var id in data) {// пробигаемся по идентификаторам
                    if (data[id].show) {// требуется вывести изменения
                        line = "[" + id + "]";
                        lines.push(line);// добавлем строку
                        line = "# " + data[id].action;
                        lines.push(line);// добавлем строку
                        for (var key in data[id].change) {// пробигаемся по ключам
                            value = data[id].change[key];// получаем значение
                            line = [// формируем строку данных
                                key, app.val.keyDelim, value
                            ].join(app.val.argDelim);
                            lines.push(line);
                        };
                        line = "";// пустая строка
                        lines.push(line);
                    };
                };
                // выводим сформированные строки
                if (lines.length) {// если есть что выводить
                    value = lines.join(app.val.lineDelim);
                    wsh.echo(value);
                };
            };
            // завершаем сценарий кодом
            wsh.quit(error);
        }
    });
})(WSH, send);
// запускаем инициализацию
send.init();