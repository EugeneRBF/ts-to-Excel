# ts-to-Excel
Converter typescript/javascript sources to/from Excel for i18n or l10n

This is small utility for converting typescript's static files to and from Excel form.
For ex. like this:
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

Its have two implementation via Excel plugin and VB standalone program.
Second one because parsing via Excel with files about 1900 lines got "Err 28 Out of stack" error
on both x32 and x64 version.<br>
When same code converted from VBA to VB works perfect w/out problems.