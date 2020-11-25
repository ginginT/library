<?php
declare (strict_types = 1);

namespace app\service;

use library\phpspreadsheet\Set;
use Monolog\Handler\StreamHandler;
use Monolog\Logger;
use think\facade\Db;

class IndexService extends \think\Service
{
    public function loopToExcel($sheet, $last_id = 0, $i = 1)
    {
        $list = Db::name('play_products')
            ->alias('pp')
            ->field('pp.id, pr.s_resources_uuid, pp.s_pro_uuid, prs.name, pps.title, pp.options')
            ->leftJoin('play_products_sub pps', 'pps.e_pro_uuid = pp.e_pro_uuid')
            ->leftJoin('play_resources pr', 'pr.e_resources_uuid = pp.e_resources_uuid')
            ->leftJoin('play_resources_sub prs', 'pr.e_resources_uuid = prs.e_resources_uuid')
            ->where([
                ['pps.lang', '=', 'zh'],
                ['prs.lang', '=', 'zh'],
                ['pp.id', '>', $last_id]
            ])
            ->limit(100)
            ->select()
            ->toArray();

        if (!$list)
            return;

        $end = array_pop($list);
        $last_id = $end['id'];

        foreach ($list as $k => $item) {
            unset($item['id']);
            $data = [];
            $options = json_decode($item['options'], true);
            if (empty($options['perBooking']) && empty($options['perPax'])) {
                unset($list[$k]);
                continue;
            }

            if (!empty($options['perBooking']))
                $data = array_merge($data, $this->handleOptions($options['perBooking'], $item['s_pro_uuid']));

            if (!empty($options['perPax']))
                $data = array_merge($data, $this->handleOptions($options['perPax'], $item['s_pro_uuid']));

            unset($item['options']);
            $item = array_values($item);

            foreach ($data as $datum) {
                array_push($item, $datum);
            }

            $list[$k] = $item;
        }

        if ($list) {
            $list = array_values($list);
            Set::setExcelBody($sheet, $list, $i);
            $i += count($list);
        }

        // 循环调用
        $this->loopToExcel($sheet, $last_id, $i);
    }

    public function handleOptions($options, $s_pro_uuid)
    {
        $data = [];

        if (!$options)
            return $data;

        foreach ($options as $option) {
            $richText = new \PhpOffice\PhpSpreadsheet\RichText\RichText();
            $this->setRichText($richText, 'uuid:', $option['uuid']);
            $this->setRichText($richText, '@@@', PHP_EOL);

            $this->setRichText($richText, 'name:', $option['name']);
            $this->setRichText($richText, '@@@', PHP_EOL);

            $this->setRichText($richText, 'nameTranslated:', $option['nameTranslated']);
            $this->setRichText($richText, '@@@', PHP_EOL);

            $this->setRichText($richText, 'description:', $option['description']);
            $this->setRichText($richText, '@@@', PHP_EOL);

            $this->setRichText($richText, 'descriptionTranslated:', $option['descriptionTranslated']);
            $this->setRichText($richText, '@@@', PHP_EOL);

            if (!empty($option['items'])) {
                foreach ($option['items'] as $item) {
                    $this->setRichText($richText, 'label:', ($item['label'] ?? '') . ', ');
                    $this->setRichText($richText, 'labelTranslated:', ($item['labelTranslated'] ?? '') . ', ');
                    $this->setRichText($richText, 'value:', ($item['value'] ?? '') . ', ');
                    $this->setRichText($richText, '@@@', PHP_EOL);
                }
            }

//            if (strlen($items) > 32767) {
//                $log = new Logger('export');
//                $path = request()->server()['DOCUMENT_ROOT']. '/../runtime/export/product/' . date('Ymd');
//                if (!is_dir($path))
//                    mkdir($path, 0777, true);
//
//                $log->pushHandler(new StreamHandler($path . '/bmg-product-to-long', Logger::WARNING));
//
//                $log->warning('Excel can store 32767 characters at most: ', [$s_pro_uuid]);
//            }

            $data[] = $richText;
        }

        return $data;
    }

    public function setRichText($richText, $text_run, $text)
    {
        $payable = $richText->createTextRun($text_run);
        $payable->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED));
        $richText->createText($text);
    }
}
