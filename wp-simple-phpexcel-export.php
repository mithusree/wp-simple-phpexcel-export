<?php
/*
  Plugin Name: Simple PHPExcel Export
  Description: Simple PHPExcel Export Plugin for WordPress
  Version: 1.0.0
  Author: Mithun
  Author URI: http://twitter.com/mithunp
 */

define("SPEE_PLUGIN_URL", WP_PLUGIN_URL.'/'.basename(dirname(__FILE__)));
define("SPEE_PLUGIN_DIR", WP_PLUGIN_DIR.'/'.basename(dirname(__FILE__)));

add_action ( 'admin_menu', 'spee_admin_menu' );

function spee_admin_menu() {
	add_menu_page ( 'PHPExcel Export', 'Export', 'manage_options', 'spee-dashboard', 'spee_dashboard' );
}

function spee_dashboard() {
	global $wpdb;
	if ( isset( $_GET['export'] )) {
		if ( file_exists(SPEE_PLUGIN_DIR . '/lib/PHPExcel.php') ) {
			
			//Include PHPExcel
			require_once (SPEE_PLUGIN_DIR . "/lib/PHPExcel.php");
			
			// Create new PHPExcel object
			$objPHPExcel = new PHPExcel();
			
			// Set document properties
			
			// Add some data
			$objPHPExcel->setActiveSheetIndex(0);
			$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
			$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Author');
			$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Date');
			$objPHPExcel->getActiveSheet()->setCellValue('D1', 'Title');
			$objPHPExcel->getActiveSheet()->setCellValue('E1', 'Status');
			$objPHPExcel->getActiveSheet()->setCellValue('F1', 'Content');
			$objPHPExcel->getActiveSheet()->setCellValue('G1', 'Comment Count');
			
			$objPHPExcel->getActiveSheet()->getStyle('A1:G1')->getFont()->setBold(true);
			$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('A:G')->setAutoSize(true);

			$query = "SELECT p.*, u.display_name
				FROM {$wpdb->prefix}posts AS p
				LEFT JOIN {$wpdb->prefix}users AS u ON p.post_author = u.ID
				WHERE p.post_type = 'post'
				ORDER BY p.ID ASC";
			
			$posts   = $wpdb->get_results($query);
			
			if ( $posts ) {
				foreach ( $posts as $i=>$post ) {
					$objPHPExcel->getActiveSheet()->setCellValue('A'.($i+2), $post->ID);
					$objPHPExcel->getActiveSheet()->setCellValue('B'.($i+2), $post->display_name);
					$objPHPExcel->getActiveSheet()->setCellValue('C'.($i+2), $post->post_date);
					$objPHPExcel->getActiveSheet()->setCellValue('D'.($i+2), $post->post_title);
					$objPHPExcel->getActiveSheet()->setCellValue('E'.($i+2), $post->post_status);
					$objPHPExcel->getActiveSheet()->setCellValue('F'.($i+2), $post->post_content);
					$objPHPExcel->getActiveSheet()->setCellValue('G'.($i+2), $post->comment_count);
				}
			}

			// Rename worksheet
			//$objPHPExcel->getActiveSheet()->setTitle('Simple');
			
			// Set active sheet index to the first sheet, so Excel opens this as the first sheet
			$objPHPExcel->setActiveSheetIndex(0);
			
			// Redirect output to a client’s web browser
			ob_clean();
			ob_start();
			switch ( $_GET['format'] ) {
				case 'csv':
					// Redirect output to a client’s web browser (CSV)
					header("Content-type: text/csv");
					header("Cache-Control: no-store, no-cache");
					header('Content-Disposition: attachment; filename="export.csv"');
					$objWriter = new PHPExcel_Writer_CSV($objPHPExcel);
					$objWriter->setDelimiter(',');
					$objWriter->setEnclosure('"');
					$objWriter->setLineEnding("\r\n");
					//$objWriter->setUseBOM(true);
					$objWriter->setSheetIndex(0);
					$objWriter->save('php://output');
					break;
				case 'xls':
					// Redirect output to a client’s web browser (Excel5)
					header('Content-Type: application/vnd.ms-excel');
					header('Content-Disposition: attachment;filename="export.xls"');
					header('Cache-Control: max-age=0');
					$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
					$objWriter->save('php://output');
					break;
				case 'xlsx':
					// Redirect output to a client’s web browser (Excel2007)
					header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
					header('Content-Disposition: attachment;filename="export.xlsx"');
					header('Cache-Control: max-age=0');
					$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
					$objWriter->save('php://output');
					break;
			}
			exit;
		}
	} 
?>
<div class="wrap">
	<h2><?php _e( "PHPExcel Export" ); ?></h2>
	<form method='get' action="admin.php?page=spee-dashboard">
		<input type="hidden" name='page' value="spee-dashboard"/>
		<input type="hidden" name='noheader' value="1"/>
		<input type="radio" name='format' id="formatCSV"  value="csv" checked="checked"/>  <label for"formatCSV">csv</label>
		<input type="radio" name='format' id="formatXLS"  value="xls"/>  <label for"formatXLS">xls</label>
		<input type="radio" name='format' id="formatXLSX" value="xlsx"/> <label for"formatXLSX">xslx</label>
		<input type="submit" name='export' id="csvExport" value="Export"/>
	</form>
	<div class="footer-credit alignright">
		<p>Thanks to awesome people at <a title="anang pratika" href="https://phpexcel.codeplex.com<" target="_blank" >PHPExcel</a>.</p>
	</div>
</div>
<?php
}
