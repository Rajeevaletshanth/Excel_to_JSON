<?php  

include_once "SimpleXLSX.php";

$question_column = 0;
$question_col_found = false;

$cont_desc_column = 0;
$cont_desc_col_found = false;

$ques_textbox_column  = 0;
$ques_textbox_col_found = false;

$row_array = array();

if(isset($_POST["upload_file"])){

    if($_FILES['file']['name']){
        $file_array = explode(".", $_FILES['file']['name']);
        $file_name = $file_array[0];
        $filename = $_FILES['file']['name'];
        $xlsx = SimpleXLSX::parse($_FILES['file']['name']);
        
        function excel_to_json($xlsx, $sheet = 0){
            global $question_column;
            global $question_col_found;
            global $cont_desc_col_found;
            global $cont_desc_column;
            global $ques_textbox_column;
            global $ques_textbox_col_found;
            
            $questions_key = ["title" => "Control/question"];
            $cont_desc_key = ["title" => "Control description"];
            $ques_textbox = ["title" => "question textbox"];
            foreach ($xlsx->rows($sheet) as $rowIndex => $row) {
            if ($rowIndex == 0) {
                foreach ($row as $colIndex => $col) {
                    if (stristr(strtolower($col), strtolower($questions_key["title"]))) {
                        $question_column = $colIndex;
                        $question_col_found = true;
                    }
                    if (stristr(strtolower($col), strtolower($cont_desc_key["title"]))) {
                        $cont_desc_column = $colIndex;
                        $cont_desc_col_found = true;
                    }
            
                    if (stristr(strtolower($col), strtolower($ques_textbox["title"]))) {
                        $ques_textbox_column = $colIndex;
                        $ques_textbox_col_found = true;
                    }
            
                }
            }
            }
            if ($question_col_found && $ques_textbox_col_found) {
            foreach ($xlsx->rows($sheet) as $rowIndex => $row) {
                if ($rowIndex > 0 && count($row) > 0) {
                    $depth = 0;
                    $row_array[] = _fill_assessment_array($row, $depth);
                }
            }
            $final_array = array();
            foreach ($row_array as $value) {
                $final_array = array_merge_recursive($final_array, $value);
            }
            return $final_array;
            }
            }
            
            function _fill_assessment_array(&$row, &$depth)
            {
            global $question_column;
            global $cont_desc_col_found;
            global $cont_desc_column;
            global $ques_textbox_column;
            global $ques_textbox_col_found;
            
            $assessment_rows = array();
            if ($row[0] == "") {
            $questions["questions"][] = array_slice($row, $question_column - $depth);
            return $questions;
            }
            if ($cont_desc_col_found && $depth == $cont_desc_column) {
            $questions["controldesc"] = strval($row[0]);
            array_shift($row);
            $depth++;
            }
            
            if ($depth == $question_column) {
            $questions["questions"][] = array_slice($row, 0, count($row));
            return $questions;
            }
            
            $control = strval($row[0]);
            array_shift($row);
            $depth++;
            $assessment_rows[$control] = _fill_assessment_array($row, $depth);
            
            return $assessment_rows;
            }
            
            function convert($array, $prefix){
            $new = array();
            $i = 1;
            foreach ($array as $controlkey => $value) {
            $prefixString = $prefix == "" ? "" : $prefix . "_";
            $controlid = $prefixString . $i;
            if (!isset($value["questions"])) {
                $new[] = array(
                    "controlid" => $controlid,
                    "controlname" => $controlkey,
                    "sub-control" => convert($value, $controlid)
                );
            } else {
                $controldesc = "";
                if (isset($value["controldesc"])) {
                    if (is_array($value["controldesc"])) {
                        $controldesc = $value["controldesc"][0];
                    } else {
                        $controldesc = $value["controldesc"];
                    }
                }
            
                $questions = array();
                foreach ($value["questions"] as $key => $quesarray) {
                    array_push($questions, array(
                        "questionid" => $key + 1,
                        "question" => array_shift($quesarray),
                        "response" => array_shift($quesarray),
                        "istextbox" => array_shift($quesarray),
                        "options" => $quesarray
                    ));
                }
            
                $new[] = array(
                    "controlid" => $controlid,
                    "controlname" => $controlkey,
                    "controldesc" => $controldesc,
                    "questions" => $questions
                );
            $i++;
            }
            return $new;
            
            }
        }

        
        $final_array = excel_to_json($xlsx);
        if (empty($final_array)) {
            echo json_encode(array("status" => "error", "response" => "Failed to read the file. The file was empty or did not have expected keywords. Please check the template file and try again!"));
        return;  
        }else{
            header('Content-disposition: attachment; filename='.$file_name.'.json');
            header('Content-type: application/json');
            echo json_encode($final_array);
            exit;
        }
    }  
}


?>

<!DOCTYPE html>
<html>
  	<head>
    	<meta name="viewport" content="width=device-width, initial-scale=1.0">
      <script src="http://code.jquery.com/jquery.js"></script>
    	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
  	</head>
  	<body>
  		<div class="container">
  			<br />
    		<br />
    		<div class="panel panel-default">
          <div class="panel-heading">
            <h3 class="panel-title">Convert XLSX to JSON</h3>
          </div>
          <div class="panel-body">
            <form method="post" enctype="multipart/form-data">
              <div class="col-md-6" align="right">Select File</div>
              <div class="col-md-6">
                <input type="file" name="file" />
              </div>
              <br /><br /><br />
              <div class="col-md-12" align="center">
                <input type="submit" name="upload_file" class="btn btn-primary" value="Upload" />
              </div>
            </form>
          </div>
        </div>
    	</div>
    	
  	</body>
</html>