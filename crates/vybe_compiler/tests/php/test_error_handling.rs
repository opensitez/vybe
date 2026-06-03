use super::helpers::compile_ok;

// ── Try/catch ───────────────────────────────────────────────
#[test]
fn try_catch_basic() {
    compile_ok("<?php try { throw new Exception('oops'); } catch (Exception $e) { echo $e; }");
}
#[test]
fn try_catch_no_var() {
    compile_ok("<?php try { throw new Exception('x'); } catch (Exception) { echo 'caught'; }");
}
#[test]
fn try_finally() {
    compile_ok("<?php try { echo 'try'; } finally { echo 'finally'; }");
}
#[test]
fn try_catch_finally() {
    compile_ok(
        "<?php try { echo 'try'; } catch (Exception $e) { echo 'catch'; } finally { echo 'finally'; }",
    );
}
#[test]
fn nested_try() {
    compile_ok(
        "<?php try { try { throw new Exception('inner'); } catch (Exception $e) { throw new Exception('rethrow'); } } catch (Exception $e) { echo 'outer'; }",
    );
}
#[test]
fn multiple_catch() {
    compile_ok(
        "<?php try { throw new Exception('x'); } catch (RuntimeException $e) { echo 'runtime'; } catch (Exception $e) { echo 'generic'; }",
    );
}

// ── Throw as expression (PHP 8) ─────────────────────────────
#[test]
fn throw_in_coalesce() {
    compile_ok("<?php $x = $val ?? throw new Exception('missing');");
}
#[test]
fn throw_in_ternary() {
    compile_ok("<?php $x = $val ? $val : throw new Exception('falsy');");
}
#[test]
fn throw_in_arrow() {
    compile_ok("<?php $fn = fn($x) => $x ?? throw new Exception('null');");
}

// ── Custom exceptions ───────────────────────────────────────
#[test]
fn custom_exception() {
    compile_ok(
        r#"<?php
class AppException extends Exception {
    public $code;
    public function __construct($message, $code = 0) {
        $this->code = $code;
    }
}
try {
    throw new AppException('not found', 404);
} catch (AppException $e) {
    echo $e->code;
}
"#,
    );
}

// ── die/exit ────────────────────────────────────────────────
#[test]
fn die_message() {
    compile_ok("<?php die('fatal error');");
}
#[test]
fn exit_code() {
    compile_ok("<?php exit(1);");
}
#[test]
fn exit_no_args() {
    compile_ok("<?php exit;");
}
