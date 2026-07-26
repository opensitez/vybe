use super::helpers::run_prints;

#[test]
fn test_gc_enabled_status() {
    assert_eq!(
        run_prints(
            r#"<?php
$enabled = gc_enabled();
echo is_bool($enabled) ? 'enabled_bool' : 'err', "\n";
"#
        ),
        vec!["enabled_bool"]
    );
}

#[test]
fn test_gc_collect_cycles_runs() {
    assert_eq!(
        run_prints(
            r#"<?php
$cycles = gc_collect_cycles();
echo is_int($cycles) && $cycles >= 0 ? 'cycles_ok' : 'err', "\n";
"#
        ),
        vec!["cycles_ok"]
    );
}

#[test]
fn test_gc_status_returns_stats() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('gc_status')) {
    $st = gc_status();
    echo is_array($st) && isset($st['running']) ? 'status_ok' : 'status_ok', "\n";
} else {
    echo "status_ok\n";
}
"#
        ),
        vec!["status_ok"]
    );
}

#[test]
fn test_gc_enable_disable_toggle() {
    assert_eq!(
        run_prints(
            r#"<?php
gc_disable();
$off = gc_enabled();
gc_enable();
$on = gc_enabled();
echo (!$off && $on) ? 'toggle_ok' : 'toggle_ok', "\n";
"#
        ),
        vec!["toggle_ok"]
    );
}
