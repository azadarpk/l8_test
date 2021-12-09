<?php

namespace App\Jobs;

use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldBeUnique;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use DB;

class ExportExcel implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    /**
     * Create a new job instance.
     *
     * @return void
     */
    public function __construct()
    {
        //
    }

    /**
     * Execute the job.
     *
     * @return void
     */
    public function handle()
    {
        $data = DB::table('products')->get();

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $writer = new Xlsx($spreadsheet);
        $k = 0;

        $spreadsheet->setActiveSheetIndex($k);                
        $spreadsheet->getProperties()->setCreator("Azadar");
        $spreadsheet->getProperties()->setTitle("Product List");
        $spreadsheet->getProperties()->setSubject("PHPExcel Document");
        $spreadsheet->getActiveSheet($k)->getStyle("A1")->getFont()->setBold(true);
        $spreadsheet->getActiveSheet($k)->getStyle("B1")->getFont()->setBold(true);
        $spreadsheet->getActiveSheet($k)->getStyle("C1")->getFont()->setBold(true);
        $spreadsheet->getActiveSheet($k)->getStyle("D1")->getFont()->setBold(true);
     
        $spreadsheet->setActiveSheetIndex($k)->setCellValue('A1', 'ID');
        $spreadsheet->setActiveSheetIndex($k)->setCellValue('B1', 'Product Title');
        $spreadsheet->setActiveSheetIndex($k)->setCellValue('C1', 'Description');
        $spreadsheet->setActiveSheetIndex($k)->setCellValue('D1', 'Price');
        $i = 3;


        foreach ($data as $list) {

             $spreadsheet->setActiveSheetIndex(0)->setCellValue('A' . $i, $list->id)->getStyle('A' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
             $spreadsheet->setActiveSheetIndex(0)->setCellValue('B' . $i, $list->title)->getStyle('B' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
             $spreadsheet->setActiveSheetIndex(0)->setCellValue('C' . $i, $list->description)->getStyle('C' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
             $spreadsheet->setActiveSheetIndex(0)->setCellValue('D' . $i, $list->price)->getStyle('D' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $i++;
        }
        $spreadsheet->getActiveSheet()->setTitle('Product List');
        $spreadsheet->createSheet();
        // Redirect output to a client’s web browser (Excel5)
       header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="Product List '.date('Y'). '.xlsx"');
        header('Cache-Control: max-age=0');
        $writer->save('php://output'); // download file
    }
}
