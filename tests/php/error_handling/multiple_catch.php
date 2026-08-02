<?php
// vybe-test: php/error_handling/multiple_catch
// origin: languages/php/tests/php/test_error_handling.rs
// vybe-test-mode: compile

try { throw new Exception('x'); } catch (RuntimeException $e) { echo 'runtime'; } catch (Exception $e) { echo 'generic'; }
