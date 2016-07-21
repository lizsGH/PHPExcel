<?php

/**
 * Class Export
 */
class Export extends Excel 
{
    /**
     * 导出 Excel
     * @param $data array 导出的数据
     * @param array $params [option]
     * 1、title：导出的 Excel 文件的标题
     * 2、columns：Excel 文件的列属性
     *  + title:第一行标题
     *  + width：对应列的宽度
     *  + field：对应列的 $data 字段
     */
    function excel_export($data, $params = array())
    {
        $excel = new PHPExcel();
        // 配置导出的 Excel 文件的属性
        $excel->getProperties()
            ->setCreator('EaseCloud')
            ->setLastModifiedBy('EaseCloud')
            ->setTitle('Title')
            ->setSubject('Subject')
            ->setDescription('Description')
            ->setKeywords('Keywords')
            ->setCategory('Category');
        // 设置 sheet
        $sheet = $excel->setActiveSheetIndex(0);

        // 设置 sheet 第一栏为标题
        $title = '';
        $field = true;
        foreach ($params['columns'] as $key => $column) {
            // 判断所有标题是否为空
            $title = $title ?: $column['title'];
            $field = $column['field'] && $field;
            $sheet->setCellValue($this->cell_column($key, 1), $column['title']);
            // 设置每列宽度
            if (!empty($column['width'])) {
                $sheet->getColumnDimension($this->cell_column($key))
                    ->setWidth($column['width']);
            }
        }
        // 如果所有标题都为空，则重置第一栏
        $row = $title ? 2 : 1;
        foreach ($data as $item) {
            // 如果所有 field 都不为空
            if ($field) {
                foreach ($params['columns'] as $key => $column) {
                    $sheet->setCellValue($this->cell_column($key, $row), $item[$column['field']]);
                }
            } else {
                foreach ($item as $k => $v) {
                    $sheet->setCellValue($this->cell_column($k, $row), $v);
                }
            }
            $row++;
        }
        $excel->getActiveSheet()->setTitle($params['title']);
        header('Content-Type: application/octet-stream');
        header('Content-Disposition: attachment;filename="' . $params['title'] . '.xls"');
        header('Cache-Control: max-age=0');
        $objWriter = PHPExcel_IOFactory::createWriter($excel, 'Excel5');
        $objWriter->save('php://output');
        exit;
    }
}