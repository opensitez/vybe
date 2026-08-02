<?php
// vybe-test: php/output_buffering/ob_start_and_ob_get_clean_preserve_newlines
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo "a\nb";
echo str_replace("\n", "|", ob_get_clean());
