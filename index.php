<h2>OpenTBS xlsx with merged cells</h2>
<a href="template.xlsx">Template file</a>
<a href="output.xlsx">Output file</a>

<?
  require_once('tbs_class.php');
  require_once('tbs_plugin_opentbs.php');

  $TBS = new clsTinyButStrong;
  $TBS->Plugin(TBS_INSTALL, OPENTBS_PLUGIN);  

  $data = array();

  for($i = 0; $i < 20; $i++)
  {
    $data[$i] = array(
      'num'  => $i + 1,
      'name' => 'Name ' . time(),
      'val'  => rand() / 1000,
    );
  }

  $template = 'template.xlsx';
  if(!file_exists($template))
  {
    echo "file $template not found";
  }
  $TBS->LoadTemplate($template);
  //$TBS->PlugIn(OPENTBS_DEBUG_INFO);
  //$TBS->PlugIn(OPENTBS_DEBUG_XML_CURRENT);  
  $TBS->PlugIn(OPENTBS_SELECT_SHEET, 1);
  $TBS->MergeBlock('a', $data);
  $TBS->PlugIn(OPENTBS_MERGE_CELLS);
  $filename = "output.xlsx";
  $TBS->Show(OPENTBS_FILE+TBS_EXIT, $filename);  
  //$TBS->Show(OPENTBS_DOWNLOAD+TBS_EXIT, $filename);  
?>
