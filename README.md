## 导入使用示例




## 导出 Excel 文件使用示例

```
$title = '抽奖序列码(' . $export_sequence['id'] . ')';
$columns = array(
    array(
        'field' => 'column1',
        'width' => 20,
    ),
    array(
        'field' => 'column2',
        'width' => 20,
    ),
    array(
        'field' => 'column3',
        'width' => 20,
    ),
    array(
        'field' => 'column4',
        'width' => 20,
    ),
    array(
        'field' => 'column5',
        'width' => 20,
    ),
);
$export_list = pdo_fetchall('...');
$rows = sizeof($export_list) / sizeof($columns);
$export_data = array();
for ($i=0; $i<$rows; $i++) {
    foreach ($columns as $key => $item) {
        $export_data[$i][$item['field']] = $export_list[$i*sizeof($columns) + $key];
    }
}
$export = new Export();
$export->excel_export($export_data, array(
    'title' => $title,
    'columns' => $columns,
));
```