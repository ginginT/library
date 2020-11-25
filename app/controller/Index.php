<?php
namespace app\controller;

use app\BaseController;
use app\service\IndexService;
use library\phpspreadsheet\Read;
use library\phpspreadsheet\Set;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Index extends BaseController
{
    public function exportProduct(IndexService $index_service)
    {
        // Excel 标题头名称
        $titles = ['资源 UUID', '产品 UUID', '玩乐资源名称', '产品名称'];
        $nums = [];

        for ($i = 0; $i < 13; $i++) {
            $titles[] = '选项 ' . ($i + 1);
        }

        foreach ($titles as $title) {
            $nums[] = 35;
        }

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        Set::setExcelTitle($sheet, $titles);
        Set::setExcelWidth($sheet, $nums);

        $index_service->loopToExcel($sheet);
        Set::setDownloadHttpHeader('export-play-product.xls');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }

    public function importProduct()
    {
        $file = request()->file('file');

        $res = Read::read($file->getPath() . DIRECTORY_SEPARATOR . $file->getFilename());

        echo '<pre>';
        var_dump($res);
    }
}
