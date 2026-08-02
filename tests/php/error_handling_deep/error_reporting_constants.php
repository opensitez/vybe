<?php
// vybe-test: php/error_handling_deep/error_reporting_constants
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

echo E_ERROR       > 0 ? 'E_ERROR ok'   : 'fail';
echo E_WARNING     > 0 ? ':E_WARNING ok'   : ':fail';
echo E_NOTICE      > 0 ? ':E_NOTICE ok'    : ':fail';
echo E_DEPRECATED  > 0 ? ':E_DEPRECATED ok': ':fail';
echo E_USER_ERROR  > 0 ? ':E_USER_ERROR ok': ':fail';
echo E_ALL         > 0 ? ':E_ALL ok'       : ':fail';
