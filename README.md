# Описание
`JScript` для отправки **переменных среды** в объект `Active Directory` по протоколу `LDAP` или в файл ярлыка в папке. Основная задача скрипта сохранить собранные данные с помощью другого скрипта [env](https://github.com/vipic-ru/env) в любой атрибут (или атрибуты) объекта `Active Directory` или на основании этих данных создать по шаблону ярлык в папке. Чтобы затем использовать этот ярлык или данные из `Active Directory` в целях системного администрирования. 

# Использование
В командной строке **Windows** введите следующую команду. Если необходимо скрыть отображение окна консоли, то вместо `cscript` можно использовать `wscript`.
```bat
cscript env.send.min.js <mode> <container> [<output>...] \\ [<input>...]
```
- `<mode>` - Режим отправки переменных (заглавное написание выполняет только эмуляцию).
    - **link** - Отправляет переменных среды в обычный ярлык.
    - **ldap** - Отправляет переменных среды в объект `Active Directory`.
- `<container>` - Путь к папке или `guid` (допускается указание пустого значения).
- `<output>` - Изменяемые свойства объекта в формате `key=value` c подстановкой переменных `%ENV%`. Первое свойство считается обязательным, т.к. по его значению осуществляется поиск объектов. Для режима **link** обязательно наличие свойств `name` и `targetpath`, а в свойстве `arguments` одинарные кавычки заменяются на двойные.
- `<input>` - Значения по умолчанию для переменных среды в формате `key=value`.

# Примеры использования
> Предполагается использовать данный скрипт совместно с другим скриптом [env](https://github.com/vipic-ru/env), поэтому сразу в примерах будут использоваться два скрипта. 

Когда компьютер в **домене**, то в **групповых политиках** при входе пользователя в компьютер, можно прописать следующий скрипт, что бы информация о компьютере и пользователе прописалась в атрибуты **описания** и **местоположения** компьютера в `Active Directory` в приделах `Organizational Unit` c **guid** `{ABCD1234-111B-14DC-ABAC-4578F1145541}`. Что бы затем быстро находить нужный компьютер пользователя или анализировать собранную информацию. Что бы узнать **guid** контейнера в `Active Directory` можно воспользоваться программой [Active Directory Explorer](https://docs.microsoft.com/ru-ru/sysinternals/downloads/adexplorer). Так же не забудьте пользователям выдать права на изменения нужных атрибутов компьютеров в соответствующем контейнере.
```bat
wscript env.min.js wscript env.send.min.js ldap {ABCD1234-111B-14DC-ABAC-4578F1145541} cn="%NET-HOST%" description="%USR-NAME-THIRD% | %USR-NAME-FIRST% %USR-NAME-SECOND% | %DEV-NAME% | %PCB-BIOS-SERIAL% | %PCB-BIOS-RELEASE-DATE% | %NET-MAC% | %DEV-BENCHMARK% | %DEV-DESCRIPTION%" location="%USR-NAME-THIRD%" \\ USR-NAME-FIRST="Terminal" USR-NAME-SECOND="login" USR-NAME-THIRD="Location" PCB-BIOS-RELEASE-DATE="XX.XX.XXXX" NET-MAC="XX:XX:XX:XX:XX:XX" 
```
Когда компьютер не в **домене** то в **планировщике задач**, можно прописать следующий скрипт, чтобы информация о компьютере и пользователе сохранялась в виде ярлычка в сетевой папке. И затем использовать эти ярлычки чтобы одним кликом разбудить `WOL` пакетом нужный компьютер и подключится к нему через **Помощник** для оказания технической поддержки. Для отправки `WOL` пакета можно использовать утилиту [Wake On Lan](https://www.depicus.com/wake-on-lan/wake-on-lan-cmd).

```bat
wscript env.min.js wscript env.send.min.js link \\server\links name="%NET-HOST% - %USR-NAME-FIRST% %USR-NAME-SECOND% ! %DEV-NAME% ! %PCB-BIOS-SERIAL% ! %PCB-BIOS-RELEASE-DATE% ! %DEV-BENCHMARK%" targetPath="%WINDIR%\System32\cmd.exe" arguments="/c wolcmd.exe %NET-MAC% 192.168.0.255 255.255.255.0 & start msra.exe /offerRA %NET-HOST%" workingDirectory="C:\Scripts" windowStyle=7 iconLocation="%WINDIR%\System32\msra.exe,0" description="%USR-NAME-THIRD%" \\ USR-NAME-FIRST="Terminal" USR-NAME-SECOND="login" PCB-BIOS-RELEASE-DATE="XX.XX.XXXX" NET-MAC="XX:XX:XX:XX:XX:XX"
```
Когда компьютер не в **домене**, но есть **административная** учётная запись от всех компьютеров, можно выполнить следующий скрипт, чтобы загрузить из `txt` файла список компьютеров, получить о них информация по сети через `WMI` и создать аналогичные ярлычки в локальной папке.
```bat
for /f "eol=; tokens=* delims=, " %%i in (list.txt) do cscript /nologo /u env.min.js \\%%i silent cscript env.send.min.js link C:\Links name="%NET-HOST% - %USR-NAME-FIRST% %USR-NAME-SECOND% ! %DEV-NAME% ! %PCB-BIOS-SERIAL% ! %PCB-BIOS-RELEASE-DATE% ! %DEV-BENCHMARK%" targetPath="%WINDIR%\System32\cmd.exe" arguments="/c wolcmd.exe %NET-MAC% 192.168.0.255 255.255.255.0 & start msra.exe /offerRA %NET-HOST%" workingDirectory="C:\Scripts" windowStyle=7 iconLocation="%WINDIR%\System32\msra.exe,0" description="%USR-NAME-THIRD%" \\ USR-NAME-FIRST="Terminal" USR-NAME-SECOND="login" PCB-BIOS-RELEASE-DATE="XX.XX.XXXX" NET-MAC="XX:XX:XX:XX:XX:XX"
```
Или можно сделать то же самое, что в предыдущем примере, но разбить всё на два этапа. Сначала получить данные с компьютеров по сети через `WMI` и сохранить их в локальной папке. А затем на основании этих данных создать аналогичные ярлычки в другой локальной папке.
```bat
for /f "eol=; tokens=* delims=, " %%i in (list.txt) do cscript /nologo /u env.min.js \\%%i > C:\Inventory\%%i.ini
for /f "eol=; tokens=* delims=, " %%i in (list.txt) do cscript /u env.min.js ini@auto silent \\ cscript env.send.min.js link C:\Links name="%NET-HOST% - %USR-NAME-FIRST% %USR-NAME-SECOND% ! %DEV-NAME% ! %PCB-BIOS-SERIAL% ! %PCB-BIOS-RELEASE-DATE% ! %DEV-BENCHMARK%" targetPath="%WINDIR%\System32\cmd.exe" arguments="/c wolcmd.exe %NET-MAC% 192.168.0.255 255.255.255.0 & start msra.exe /offerRA %NET-HOST%" workingDirectory="C:\Scripts" windowStyle=7 iconLocation="%WINDIR%\System32\msra.exe,0" description="%USR-NAME-THIRD%" \\ USR-NAME-FIRST="Terminal" USR-NAME-SECOND="login" PCB-BIOS-RELEASE-DATE="XX.XX.XXXX" NET-MAC="XX:XX:XX:XX:XX:XX" < C:\Inventory\%%i.ini
```