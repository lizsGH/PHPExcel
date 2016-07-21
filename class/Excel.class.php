<?php

/**
 * Class Excel
 */
class Excel
{
    function __construct() {
        if (PHP_SAPI == 'cli') {
            die("This example should only be run from a web browser!");
        }
        // 导入 PHPExcel 核心类
        require_once "../inc/PHPExcel.php";
    }

    /**
     * get cell column name
     * @param int $key
     * @param int $row
     * @return bool|string
     */
    function cell_column($key = 0, $row = 0) {
        $row = $row ?: '';
        if ($key >= 0 && $key < 26) {
            return chr(65 + $key) . $row;
        } elseif ($key >= 26) {
            return chr(intval($key/26) + 64) . chr($key%26 + 65) . $row;
        } else {
            return false;
        }
    }


}