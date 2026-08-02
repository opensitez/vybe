<?php
// vybe-test: php/output_buffering/ob_callback_can_strip_tags
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(fn(string $buf): string => strip_tags($buf));
echo '<b>ok</b> <i>yes</i>';
ob_end_flush();
