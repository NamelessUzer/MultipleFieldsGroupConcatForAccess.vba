# MultipleFieldsGroupConcatForAccess.vba
这是一个用在Access中的GroupConcat函数，支持多字段分组、支持排序。

## 用法
```sql
SELECT 字段1, 字段2, 字段3, GroupConcat(BuildQuery("字段1", [字段1], "字段2", [字段2]),"拼接字段","tableName","、","排序字段") AS ConcatedFields
FROM tableName
GROUP BY 字段1, 字段2, 字段3
ORDER BY 字段1, 字段2, 字段3;
```
参数说明：
第一个参数是函数BuildQuery()，该函数支持可变数量参数，参数每两个为一对，每一对的前一个是用引号包围起来的字段名，后一个是用“[]”包围起来的字段名，支持多字段，每一对参数都表示一个用来分组的字段；
第二个参数为待拼接的字段名；
第三个参数是要查询的数据表名；
第四个参数是拼接所用到的字符或字符串，默认值为“; ”；
第五个参数为排序字段。
