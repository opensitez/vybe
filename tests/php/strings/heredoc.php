<?php
// vybe-test: php/strings/heredoc
// origin: languages/php/tests/php/test_strings.rs
// vybe-test-mode: compile

$x = <<<EOT
Hello World
EOT;
echo $x;
