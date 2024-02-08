<#list data as row>
    <#list row?keys as key>
        ${row[key]}<#if key_has_next>,</#if>
    </#list>
    <#if row_has_next>
        ${'\n'}
    </#if>
</#list>
