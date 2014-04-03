<?php

	require_once dirname(__FILE__)."/lib/pclzip.lib.php";

	class phpdocx{
		
		private $template;
		private $tmpDir = "/tmp/phpdocx"; // must be writable
		private $assigned_field = array();
		private $assigned_block = array();
		private $assigned_nested_block = array();
		private $assigned_header_field = array();
		private $assigned_footer_field = array();
		private $block_content = array();
		private $block_count = array();
		private $images = array();
		private $nested_block_count = array();

		public function __construct($template){
		
			if(file_exists($template)){
				$this->template = $template;
			} else {
				throw new Exception("The template ".$template." was not found !");
			}
			
		
		}
		
		/* the ref will be used to assign this image later 
		 * You can use whatever you want
		 * */
		public function addImage($ref,$file){
			if(file_exists($file)){
				if(!array_key_exists($ref,$this->images)){
					$this->images[$ref] = $file;
				} else {
					throw new Exception("the ref $ref allready exists");
				}
			} else {
				throw new Exception($file.' does not exist');
			}
		}
		
		public function assign($field,$value){
			$this->assigned_field[$field] = $this->filter($value);
		}
		
		public function assignToHeader($field,$value){
			$this->assigned_header_field[$field] = $this->filter($value);
		}
		
		public function assignToFooter($field,$value){
			$this->assigned_footer_field[$field] = $this->filter($value);
		}
		
		public function assignBlock($blockname,$values){
			$this->assigned_block[$blockname] = $values;
		}
		
		public function assignNestedBlock($blockname,$values,$parent){
			array_push($this->assigned_nested_block,array("block"=>$blockname,"values"=>$values,"parent"=>$parent));
		}
		
		public function download($name = null){
			
			$tmp_filename = $this->tmpDir."/".uniqid(true).".docx";
			
			if(is_null($name)){
				$name = basename($tmp_filename);
			}
			
			$this->save($tmp_filename);
			
			header('Content-Description: File Transfer');
			header('Content-Type: application/octet-stream');
			header('Content-Disposition: attachment; filename="' . $name . '"');
			header('Content-Transfer-Encoding: binary');
			header('Expires: 0');
			header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
			header('Pragma: public');
			header('Content-Length: ' . filesize($tmp_filename));
			ob_clean();
			flush();
			readfile($tmp_filename);
			exit;
		}
		
		
		private function saveMainDocument(){
			$this->content = file_get_contents($this->tmpDir."/word/document.xml");
			$this->clean();
			
			$this->content = $this->parseBlocks($this->content);
			
						
			foreach($this->assigned_field as $field => $value){
				$this->content = str_replace($field,$value,$this->content);
			}
		
			if(count($this->assigned_block)>0){
				foreach($this->assigned_block as $block => $values){
					
					foreach($values as $value){
						$this->addBlock($block,$value);
					}
					
				}
			}
			
			if(count($this->assigned_nested_block)>0){
				foreach($this->assigned_nested_block as $array){
					$this->addNestedBlock($array['block'],$array['values'],$array['parent']);
				}
			}
			
			file_put_contents($this->tmpDir."/word/document.xml",$this->content);
		}
		
		private function saveHeader(){
      foreach ($this->getFilesStartingWith('header') as $header_file) {
      
        $this->headerContent = file_get_contents($this->tmpDir."/word/" . $header_file);

        foreach($this->assigned_header_field as $field => $value){
          $this->headerContent = str_replace($field,$value,$this->headerContent);
        }

        file_put_contents($this->tmpDir."/word/" . $header_file, $this->headerContent);
      }
		}
		
		
		private function saveFooter(){
      foreach ($this->getFilesStartingWith('footer') as $footer_file) {
        $this->footerContent = file_get_contents($this->tmpDir."/word/" . $footer_file);

        foreach($this->assigned_footer_field as $field => $value){
          $this->footerContent = str_replace($field,$value,$this->footerContent);
        }

        file_put_contents($this->tmpDir."/word/" . $footer_file, $this->footerContent);
      }
		}
		
		//assigned_header_field
		public function save($outputFile){
			
			$this->extract();
			
			$this->processImages();
			
			$this->saveMainDocument();
			
			
			if(count($this->assigned_header_field) > 0){
				$this->saveHeader();
			}
			
			if(count($this->assigned_footer_field) > 0){
				$this->saveFooter();
			}
			
			$this->compact($outputFile);
		}
		
		private function processImages(){

			if(count($this->images)>0){
				
				if(!is_dir($this->tmpDir."/word/media")){
					mkdir($this->tmpDir."/word/media");
				}
				
				$relationships = file_get_contents($this->tmpDir."/word/_rels/document.xml.rels");
				
				foreach($this->images as $ref => $file){
					$xml = '<Relationship Id="phpdocx_'.$ref.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/'.basename($file).'" />';
					copy($file,$this->tmpDir."/word/media/".basename($file));
					$relationships = str_replace('ships">','ships">'.$xml,$relationships);
				}
				
				file_put_contents($this->tmpDir."/word/_rels/document.xml.rels",$relationships);
				
			}
			
		}
		
		private function extract(){
			
			if(file_exists($this->tmpDir) && is_dir($this->tmpDir)){
				//clean up of the tmp dir
				$this->rrmdir($this->tmpDir);
			}
			
			mkdir($this->tmpDir);
			
			$archive = new PclZip($this->template);
			$archive->extract(PCLZIP_OPT_PATH, $this->tmpDir);
			
		}
		
		private function compact($output){
		
		
			$archive = new PclZip($output);
			$archive->create($this->tmpDir,PCLZIP_OPT_REMOVE_PATH,$this->tmpDir);
		}
		
		private function rrmdir($dir) {
		   if (is_dir($dir)) {
			 $objects = scandir($dir);
			 foreach ($objects as $object) {
			   if ($object != "." && $object != "..") {
				 if (filetype($dir."/".$object) == "dir") $this->rrmdir($dir."/".$object); else unlink($dir."/".$object);
			   }
			 }
			 reset($objects);
			 rmdir($dir);
		   }
		}
		
		private function clean(){
			$this->content = str_replace('<w:lastRenderedPageBreak/>', '', $this->content); // faster
			$this->cleanTag(array('<w:proofErr', '<w:noProof', '<w:lang', '<w:lastRenderedPageBreak'));
			$this->cleanRsID();
			$this->cleanDuplicatedLayout();
		}
		
		
		private function parseBlocks($txt){
		
			$matches = array();
			$ret = $txt;
			
			preg_match_all('/\[start (\w+)\].*?\[end \1\]/s',$txt,$matches);
			
			if(count($matches[1])>0){
				foreach($matches[1] as $block){
					$ret = $this->parseBlock($block,$ret);
				}
			}
			
			return $ret;
		
		}
		
		private function parseBlock($name,$txt){
			
			/* we strip the block markup */
			$previous_pos = $this->getPreviousPosOf("start ".$name,"<w:p ",$txt);
			$end_pos = $this->getNextPosOf("start ".$name,":p>",$txt) + 3;
			
			$txt = str_replace(substr($txt,$previous_pos,$end_pos-$previous_pos),"<!-- start ".$name." -->",$txt);
			
			$previous_pos = $this->getPreviousPosOf("end ".$name,"<w:p ",$txt);
			$end_pos = $this->getNextPosOf("end ".$name,":p>",$txt) + 3;
			
			$txt = str_replace(substr($txt,$previous_pos,$end_pos-$previous_pos),"<!-- end ".$name." -->",$txt);
			
			/* we save the template content for the block */
			$block = preg_match("`<!-- start ".$name." -->(.*)<!-- end ".$name." -->`",$txt,$matches);
			
			if(array_key_exists(1,$matches) > 0){
				$this->block_content[$name] = $this->parseBlocks($matches[1]);
			}
			
		
			/* we remove the template content from the doc */
			$txt = preg_replace('`<!-- start '.$name.' -->(.*)<!-- end '.$name.' -->`','<!-- start '.$name.' --><!-- end '.$name.' -->',$txt);

			return $txt;
		
		}
		
		private function addBlock($blockname,$values){
			
			$block = $this->block_content[$blockname];
			
			if(array_key_exists($blockname,$this->block_count)){
				$this->block_count[$blockname] = $this->block_count[$blockname] + 1;
			} else {
				$this->block_count[$blockname] = 1;
			}
			
			foreach($values as $id => $val){
				$block = str_replace($id,$this->filter($val),$block);
			}
			
			$this->content = str_replace("<!-- end ".$blockname." -->","<!-- block_".$blockname."_".$this->block_count[$blockname]." -->".$block."<!-- end_block_".$blockname."_".$this->block_count[$blockname]." --><!-- end ".$blockname." -->",$this->content);
			
		}
		
		/**
		 * parent is an array of parents with the blockname as id and index as key
		 * For example :
		 * array("experience"=>1,"project"=>5);
		 */
		public function addNestedBlock($blockname,$values,$parent){
			
			if(is_array($parent) && count($parent)>0){
			
				$block = "";
				$regex = '`(.*)`';
				
				$link_nested_count = array();
				
				foreach($parent as $id => $node){
					
					if($regex == "`(.*)`"){
						$regex = str_replace("(.*)","<!-- block_".$id."_".$node." -->(.*)<!-- end_block_".$id."_".$node." -->",$regex);
					} else {
						$regex = str_replace("(.*)",".*<!-- block_".$id."_".$node." -->(.*)<!-- end_block_".$id."_".$node." -->.*",$regex);
					}
					
					array_push($link_nested_count,$id.$node);
				}
				
				$idnested = implode("_",$link_nested_count)."_".$blockname;
				
				if(array_key_exists($idnested,$this->nested_block_count)){
					$current_index = $this->nested_block_count[$idnested] + 1;
					$this->nested_block_count[$idnested]++;
				} else {
					$this->nested_block_count[$idnested] = 1;
					$current_index = 1;
				}
				
				$block_content = $this->block_content[$blockname];
				
				foreach($values as $row){
					$current_block = $block_content;
					foreach($row as $id => $val){
						$current_block = str_replace($id,$this->filter($val),$current_block);
					}
					$block .= $current_block;
				}
				
				preg_match($regex,$this->content,$matches);
				
				$new = str_replace("<!-- end ".$blockname." -->","<!-- block_".$blockname."_".$current_index." -->".$block."<!-- end_block_".$blockname."_".$current_index." --><!-- end ".$blockname." -->",$matches[0]);
				
				$this->content = preg_replace($regex,$new,$this->content);
			
			} else {
				
				die("Parent cannot be empty");
				
			}
			
		}
		
		private function filter($value){
			return str_replace("&","&amp;",$value);
		}
		
		private function cleanRsID() {
		/* From TBS script
		 * Delete XML attributes relative to log of user modifications. Returns the number of deleted attributes.
		In order to insert such information, MsWord do split TBS tags with XML elements.
		After such attributes are deleted, we can concatenate duplicated XML elements. */

			$rs_lst = array('w:rsidR', 'w:rsidRPr');

			$nbr_del = 0;
			foreach ($rs_lst as $rs) {

				$rs_att = ' '.$rs.'="';
				$rs_len = strlen($rs_att);

				$p = 0;
				while ($p!==false) {
					// search the attribute
					$ok = false;
					$p = strpos($this->content, $rs_att, $p);
					if ($p!==false) {
						// attribute found, now seach tag bounds
						$po = strpos($this->content, '<', $p);
						$pc = strpos($this->content, '>', $p);
						if ( ($pc!==false) && ($po!==false) && ($pc<$po) ) { // means that the attribute is actually inside a tag
							$p2 = strpos($this->content, '"', $p+$rs_len); // position of the delimiter that closes the attribute's value
							if ( ($p2!==false) && ($p2<$pc) ) {
								// delete the attribute
								$this->content = substr_replace($this->content, '', $p, $p2 -$p +1);
								$ok = true;
								$nbr_del++;
							}
						}
						if (!$ok) $p = $p + $rs_len;
					}
				}

			}

			// delete empty tags
			$this->content = str_replace('<w:rPr></w:rPr>', '', $this->content);
			$this->content = str_replace('<w:pPr></w:pPr>', '', $this->content);

			return $nbr_del;

		}
		
		private function cleanDuplicatedLayout() {
		// Return the number of deleted dublicates
		
			$wro = '<w:r';
			$wro_len = strlen($wro);

			$wrc = '</w:r';
			$wrc_len = strlen($wrc);

			$wto = '<w:t';
			$wto_len = strlen($wto);

			$wtc = '</w:t';
			$wtc_len = strlen($wtc);

			$nbr = 0;
			$wro_p = 0;
			while ( ($wro_p=$this->foundTag($this->content, $wro, $wro_p))!==false ) {
				$wto_p = $this->foundTag($this->content,$wto,$wro_p); if ($wto_p===false) return false; // error in the structure of the <w:r> element
				$first = true;
				do {
					$ok = false;
					$wtc_p = $this->foundTag($this->content,$wtc,$wto_p); if ($wtc_p===false) return false; // error in the structure of the <w:r> element
					$wrc_p = $this->foundTag($this->content,$wrc,$wro_p); if ($wrc_p===false) return false; // error in the structure of the <w:r> element
					if ( ($wto_p<$wrc_p) && ($wtc_p<$wrc_p) ) { // if the found <w:t> is actually included in the <w:r> element
						if ($first) {
							$superflous = '</w:t></w:r>'.substr($this->content, $wro_p, ($wto_p+$wto_len)-$wro_p); // should be like: '</w:t></w:r><w:r>....<w:t'
							$superflous_len = strlen($superflous);
							$first = false;
						}
						$x = substr($this->content, $wtc_p+$superflous_len,1);
						if ( (substr($this->content, $wtc_p, $superflous_len)===$superflous) && (($x===' ') || ($x==='>')) ) {
							// if the <w:r> layout is the same same the next <w:r>, then we join it
							$p_end = strpos($this->content, '>', $wtc_p+$superflous_len); //
							if ($p_end===false) return false; // error in the structure of the <w:t> tag
							$this->content = substr_replace($this->content, '', $wtc_p, $p_end-$wtc_p+1);
							$nbr++;
							$ok = true;
						}
					}
				} while ($ok);

				$wro_p = $wro_p + $wro_len;

			}

			return $nbr; // number of replacements

		}
		
		private function foundTag($Txt, $Tag, $PosBeg) {
		// Found the next tag of the asked type. (Not specific to MsWord, works for any XML)
			$len = strlen($Tag);
			$p = $PosBeg;
			while ($p!==false) {
				$p = strpos($Txt, $Tag, $p);
				if ($p===false) return false;
				$x = substr($Txt, $p+$len, 1);
				if (($x===' ') || ($x==='/') || ($x==='>') ) {
					return $p;
				} else {
					$p = $p+$len;
				}
			}
			return false;
		}
		
		private function cleanTag($TagLst) {
		// Delete all tags of the types listed in the list. (Not specific to MsWord, works for any XML)
			$nbr_del = 0;
			foreach ($TagLst as $tag) {
				$p = 0;
				while (($p=$this->foundTag($this->content, $tag, $p))!==false) {
					// get the end of the tag
					$pe = strpos($this->content, '>', $p);
					if ($pe===false) return false; // error in the XML formating
					// delete the tag
					$this->content = substr_replace($this->content, '', $p, $pe-$p+1);
				} 
			}
			return $nbr_del;
		}
		
		private function getNextPosOf($start_string,$needle,$txt){
			$current_pos = strpos($txt,$start_string);
			$len = strlen($needle);
			$not_found = true;
			while($not_found && $current_pos <= strlen($this->content)){
				if(substr($txt,$current_pos,$len) == $needle){
					return $current_pos;
				} else {
					$current_pos = $current_pos + 1;
				}
			}
			
			return 0;
		}
		
		private function getPreviousPosOf($start_string,$needle,$txt){
			$current_pos = strpos($txt,$start_string);
			$len = strlen($needle);
			$not_found = true;
			while($not_found && $current_pos >= 0){
				if(substr($txt,$current_pos,$len) == $needle){
					return $current_pos;
				} else {
					$current_pos = $current_pos - 1;
				}
			}
			
			return 0;
		}

    /**
     * Scan directory for files starting with $file_name
     */
		private function getFilesStartingWith($file_name) {
      $dir = $this->tmpDir . "/word/";
      static $files;
      if (!isset($files)) {
        $files = scandir($dir);
      }
      
      $found_files = array();
      foreach ($files as $file) {
        if (is_file($dir . $file)) {
          if (strpos($file, $file_name) !== FALSE) {
            $found_files[] = $file;
          }
        }
      }
      
      return $found_files;
    }
	}

?>
