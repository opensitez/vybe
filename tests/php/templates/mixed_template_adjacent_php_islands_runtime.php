<?php
// vybe-test: php/templates/mixed_template_adjacent_php_islands_runtime
// origin: languages/php/tests/php/test_templates.rs

$i = 1; $files = [null, ["isBack" => true]]; ?><tr class="snF <?php echo ($i%2==0) ? "snEven" : "snOdd"?><?php echo (isset($files[$i]["isBack"]) && $files[$i]["isBack"]) ? ' snBack' : '';?>"></tr>
