<?php
// vybe-test: php/output_buffering/ob_get_clean_drops_nested_inner_when_called_in_inner
// origin: languages/php/tests/php/test_output_buffering.rs

echo 'base';
ob_start();
echo 'one';
ob_start();
echo 'two';
echo '|' . ob_get_clean();
echo ob_get_clean();
