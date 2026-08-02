<?php
// vybe-test: php/php7/php71_multi_catch
// origin: languages/php/tests/php/test_php7.rs
// vybe-test-mode: compile

try { } catch (TypeError | ValueError $e) { echo $e; }
