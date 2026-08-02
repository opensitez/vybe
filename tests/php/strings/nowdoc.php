<?php
// vybe-test: php/strings/nowdoc
// origin: languages/php/tests/php/test_strings.rs
// vybe-test-mode: compile

$x = <<<'EOT'
No $interpolation
EOT;
echo $x;
