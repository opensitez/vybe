<?php
// vybe-test: php/templates/mixed_template_attribute_echo_runtime
// origin: languages/php/tests/php/test_templates.rs

$cols = [1, 2]; ?><td colspan="<?php echo count($cols)?>">ok</td>
