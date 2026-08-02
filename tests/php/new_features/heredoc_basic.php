<?php
// vybe-test: php/new_features/heredoc_basic
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

$x = <<<EOT
Hello World
EOT;
echo $x;
