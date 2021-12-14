Основные возможности этого проекта состоят в том, чтобы взять исходный код на javascript, в котором обычно хранятся строки для локализации какого нибудь проекта SPA,
в котором например используется библиотека angular-translate со статическими файлами, такого вида:
<pre><code>
export const translations = {};<br>
translations['ru'] = {<blockquote>
    DELETE_LIST: 'Удалить список',<br>
    CHANGE_LIST: 'Изменить список',<br>
    LIST: {<blockquote>
        UPDATE_TYPE: 'Изменить тип',<br>
        NEW: 'Новый список',<br>
        EDIT: 'Изменить',<br>
        VIEW: 'Просмотр',<br>
    },</blockquote>
    ADD: 'Добавить',<br>
    ADD_TO_START: 'Добавить в начало',<br>
    ADD_TO_END: 'Добавить в конец',<br>
};</blockquote>
translations['en'] = {<blockquote>
    DELETE_LIST: 'Delete list',<br>
    CHANGE_LIST: 'Change list',<br>
    LIST: {<blockquote>
        UPDATE_TYPE: 'Update type',<br>
        NEW: 'New list',<br>
        EDIT: 'Edit',<br>
        VIEW: 'View',<br>
    },</blockquote>
    ADD: 'Add',<br>
    ADD_TO_START: 'Add to start',<br>
    ADD_TO_END: 'Add to end',<br>
};</blockquote>
</code></pre>

И конвертировать его в Excel таблицу такого вида:

<table><tr><th>KEY</th><th>ru</th><th>en</th></tr>
<tr><td>ADD</td><td>Добавить</td><td>Add</td></tr>
<tr><td>ADD_TO_END</td><td>Добавить в конец</td><td>Add to end</td></tr>
<tr><td>ADD_TO_START</td><td>Добавить в начало</td><td>Add to start</td></tr>
<tr><td>CHANGE_LIST</td><td>Изменить список</td><td>Change list</td></tr>
<tr><td>DELETE_LIST</td><td>Удалить список</td><td>Delete list</td></tr>
<tr><td>LIST.EDIT</td><td>Изменить</td><td>Edit</td></tr>
<tr><td>LIST.NEW</td><td>Новый список</td><td>New list</td></tr>
<tr><td>LIST.UPDATE_TYPE</td><td>Изменить тип</td><td>Update type</td></tr>
<tr><td>LIST.VIEW</td><td>Просмотр</td><td>View</td></tr>
</table>

Строки в таблице отсортированы, ключи для вложенных фраз конкатенируются через точку, т.е. приводятся к виду, в котором они используются в проекте и для сохранения уникальности.

Такой вид обычно подходит для переводчиков, всегда можно добавить новую колонку для нового языка, можно добавить примечания или пометки для переводчика.
Программа может сделать обратную конвертацию в файл javascript. Но потребуется поправить код программы, чтобы изменить шаблон, который будет выведен перед значением json.<br>
Таких шаблонов 2шт - 1й перед всем текстом, 2й перед значением для каждой локали<br>

Также программа позволяет выбрать сразу несколько файлов, что удобно, если переводы лежат для каждого языка в отдельном файле.
Также игнорируются строки вида
<pre><code>
import 'some-package-name'; 
</code></pre>
а также игнорируются inline коментарии в /* ... */ и строчные, типа //... в любом месте исходного кода.


