<?php

require("phpDocx.php");

$phpdocx = new phpdocx("template.docx");
$phpdocx->assign("#TITLE1#","Pet shop list"); // basic field mapping

$phpdocx->assignBlock("members",array(array("#NAME#"=>"John","#SURNAME#"=>"DOE"),array("#NAME#"=>"Jane","#SURNAME#"=>"DOE"))); // this would replicate two members block with the associated values

$phpdocx->assignNestedBlock("pets",array(array("#PETNAME#"=>"Rex")),array("members"=>1)); // would create a block pets for john doe with the name rex
$phpdocx->assignNestedBlock("pets",array(array("#PETNAME#"=>"Rox")),array("members"=>2)); // would create a block pets for jane doe with the name rox

$phpdocx->assignNestedBlock("toys",array(array("#TOYNAME#"=>"Ball"),array("#TOYNAME#"=>"Frisbee"),array("#TOYNAME#"=>"Box")),array("members"=>1,"pets"=>1)); // would create a block toy for rex
$phpdocx->assignNestedBlock("toys",array(array("#TOYNAME#"=>"Frisbee")),array("members"=>2,"pets"=>1)); // would create a block toy for rox

$phpdocx->save("pets.docx");
?>
