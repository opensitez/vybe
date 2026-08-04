<?php
// vybe-test: php/templates/mixed_template_inline_if_attribute_runtime
// origin: languages/php/tests/php/test_templates.rs

$w = 120; ?><td<?php if ($w>0) echo " style=\"width:".$w."px;\"";?>>x</td>
