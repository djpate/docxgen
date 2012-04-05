PHPDOCX
=======

Fork
----
This is a fork from the original https://github.com/djpate/docxgen
Just needed to also change some fields in the header and footer so I created the necessary methods to do so.

Note: This are very basic methods, it assumes there's only 1 header (header1.xml) and 1 footer  (footer1.xml).

I'm no expert in the docx format so I just replaced the tag with the value, no xml iteration or funky business here :)

The example should be self explanatory.



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

	require("phpDocx.php");
	$phpdocx = new phpdocx("mytemplate.docx");

How to assign values
--------------------

	$phpdocx->assign("#TITLE1#","Hello !"); // basic field mapping

	$phpdocx->assignBlock("members",array(array("#NAME#"=>"John","#SURNAME#"=>"DOE"),array("#NAME#"=>"Jane","#SURNAME#"=>"DOE"))); // this would replicate two members block with the associated values

	$phpdocx->assignNestedBlock("pets",array(array("#PETNAME#"=>"Rex")),array("members"=>1)); // would create a block pets for john doe with the name rex
	$phpdocx->assignNestedBlock("pets",array(array("#PETNAME#"=>"Rox")),array("members"=>2)); // would create a block pets for jane doe with the name rox

	$phpdocx->assignNestedBlock("toys",array(array("#TOYNAME#"=>"Ball")),array("members"=>1,"pets"=>1)); // would create a block toy for rex
	$phpdocx->assignNestedBlock("toys",array(array("#TOYNAME#"=>"Frisbee")),array("members"=>2,"pets"=>1)); // would create a block toy for rox

How to save
-----------
you can either save the generated doc to a file using the save method

	$phpdocx->save("somefile.docx");
	
Or you can force the download of the file with the download method
	
	$phpdocx->download("somefile.docx");
	
If you do not specify a filename in the download method a generated filename will be used
	
	
More info
---------


### Why this pclzip ?


I'm using pclzip for the zip process because the zip utility provided with php can cause issue with office


### What's the licence ?

GPL

### Anything else ?

I'm using three function from the TBS library so congrats to them




