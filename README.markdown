PHPDOCX
=======

Features
--------

+ Create valid docx based on a template
+ Nested blocks on infinite levels
    
How to create your template
---------------------------

Simply open up word 2007+

If you want to map a single field you can just use #NAME# but you could use anything you like since it's just a search & replace

To create a block

[start blockname]
your content
[start somenestedblock]
[end somenestedblock]
[end blockname]

+ blockname should be unique
+ blockname has to match \w+ reg ex (a-zA-Z_)
+ fields mapped in block has to be unique

Please check out full template in the repo


How to setup
------------

require("phpdocx.php");

$phpdocx = new phpdocx("mytemplate.docx");

How to assign values
--------------------

$phpdocx->assign("#TITLE1#","Hello !"); // basic field mapping

$phpdocx->assignBlock("members",array(array("#NAME#"=>"John","#SURNAME#"=>"DOE"),array("#NAME#"=>"Jane","#SURNAME#"=>"DOE"))); // this would replicate two members block with the associated values

$phpdocx->assignNestedBlock("pets",array("#PETNAME#"=>"Rex"),array("members"=>1)); // would create a block pets for john doe with the name rex
$phpdocx->assignNestedBlock("pets",array("#PETNAME#"=>"Rox"),array("members"=>2)); // would create a block pets for jane doe with the name rox

$phpdocx->assignNestedBlock("toys",array("#TOYNAME#"=>"Ball"),array("members"=>1,"pets"=>1)); // would create a block toy for rex
$phpdocx->assignNestedBlock("toys",array("#TOYNAME#"=>"Frisbee"),array("members"=>2,"pets"=>1)); // would create a block toy for rox





