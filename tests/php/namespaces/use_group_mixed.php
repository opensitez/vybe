<?php
// vybe-test: php/namespaces/use_group_mixed
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Util;
class Logger { public function log(string $m): void { echo $m; } }
function format(string $s): string { return "[$s]"; }
const LOG_LEVEL = 'info';

namespace App;
use Util\{Logger, function format, const LOG_LEVEL};
$log = new Logger();
$log->log(format(LOG_LEVEL));
