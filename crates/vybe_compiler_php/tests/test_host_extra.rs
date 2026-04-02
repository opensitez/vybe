mod helpers;
use helpers::compile_ok;

// ── Convert / Math extras ───────────────────────────────────
#[test] fn dechex() { compile_ok("<?php echo dechex(255);"); }
#[test] fn decoct() { compile_ok("<?php echo decoct(8);"); }
#[test] fn is_finite() { compile_ok("<?php echo is_finite(1.0);"); }
#[test] fn is_nan() { compile_ok("<?php echo is_nan(NAN);"); }
#[test] fn hypot() { compile_ok("<?php echo hypot(3, 4);"); }
#[test] fn intdiv() { compile_ok("<?php echo intdiv(7, 2);"); }
#[test] fn fmod() { compile_ok("<?php echo fmod(10.5, 3.2);"); }
#[test] fn fdiv() { compile_ok("<?php echo fdiv(10, 3);"); }
#[test] fn pi_const() { compile_ok("<?php echo pi();"); }

// ── Network ─────────────────────────────────────────────────
#[test] fn gethostbyname() { compile_ok("<?php $ip = gethostbyname('localhost');"); }

// ── CLI / IO ────────────────────────────────────────────────
#[test] fn readline() { compile_ok("<?php $input = readline();"); }
#[test] fn error_log() { compile_ok("<?php error_log('something went wrong');"); }
#[test] fn trigger_error() { compile_ok("<?php trigger_error('warning', E_USER_WARNING);"); }
#[test] fn phpinfo() { compile_ok("<?php phpinfo();"); }
#[test] fn get_current_user() { compile_ok("<?php echo get_current_user();"); }

// ── Filesystem extras ───────────────────────────────────────
#[test] fn stat_file() { compile_ok("<?php $info = stat('/tmp/test.txt');"); }
#[test] fn readdir() { compile_ok("<?php $entries = readdir('/tmp');"); }

// ── HTTP extended ───────────────────────────────────────────
#[test] fn fetch_api() { compile_ok("<?php $resp = fetch('https://api.example.com/data');"); }

// ── DateTime ────────────────────────────────────────────────
#[test] fn datetime_now() { compile_ok("<?php $now = new DateTime(); echo $now->format('Y-m-d');"); }
#[test] fn datetime_parse() { compile_ok("<?php $dt = new DateTime('2024-01-15'); echo $dt->format('Y-m-d');"); }
#[test] fn datetime_immutable() { compile_ok("<?php $dt = new DateTimeImmutable();"); }
#[test] fn datetime_timestamp() { compile_ok("<?php $dt = new DateTime(); $ts = $dt->getTimestamp();"); }

// ── SplStack ────────────────────────────────────────────────
#[test] fn spl_stack() { compile_ok(r#"<?php
$stack = new SplStack();
$stack->push('a');
$stack->push('b');
$stack->push('c');
$top = $stack->pop();
echo $top;
"#); }

// ── SplQueue ────────────────────────────────────────────────
#[test] fn spl_queue() { compile_ok(r#"<?php
$queue = new SplQueue();
$queue->enqueue('first');
$queue->enqueue('second');
$item = $queue->dequeue();
echo $item;
$next = $queue->peek();
"#); }

// ── Real-world combined ─────────────────────────────────────
#[test] fn datetime_db_pattern() { compile_ok(r#"<?php
$pdo = new PDO('sqlite:app.db');
$now = new DateTime();
$timestamp = $now->format('Y-m-d H:i:s');
$pdo->exec("INSERT INTO logs (created_at) VALUES ('" . $timestamp . "')");
"#); }

#[test] fn cli_app() { compile_ok(r#"<?php
echo "Running on: " . php_uname() . "\n";
echo "User: " . get_current_user() . "\n";
echo "CWD: " . getcwd() . "\n";
echo "PHP: " . phpversion() . "\n";
$name = readline();
echo "Hello, " . $name . "!\n";
"#); }

#[test] fn hex_color() { compile_ok(r#"<?php
function hexColor($r, $g, $b) {
    return '#' . str_pad(dechex($r), 2, '0') . str_pad(dechex($g), 2, '0') . str_pad(dechex($b), 2, '0');
}
echo hexColor(255, 128, 0);
"#); }

#[test] fn stack_based_calculator() { compile_ok(r#"<?php
$stack = new SplStack();
$tokens = explode(' ', '3 4 + 2 *');
foreach ($tokens as $token) {
    if (is_numeric($token)) {
        $stack->push(intval($token));
    } else {
        $b = $stack->pop();
        $a = $stack->pop();
        $result = match($token) {
            '+' => $a + $b,
            '-' => $a - $b,
            '*' => $a * $b,
            '/' => $a / $b,
            default => 0
        };
        $stack->push($result);
    }
}
echo $stack->pop();
"#); }
