use super::helpers::run_prints;

// ── PHP version constants ─────────────────────────────────────

#[test]
fn php_version_defined() {
    assert_eq!(
        run_prints(r#"<?php echo PHP_MAJOR_VERSION >= 8 ? 'php8+' : 'old'; "#),
        vec!["php8+"]
    );
}
#[test]
fn php_version_id() {
    assert_eq!(
        run_prints(r#"<?php echo PHP_VERSION_ID >= 80000 ? 'ok' : 'old'; "#),
        vec!["ok"]
    );
}
#[test]
fn php_os_family_defined() {
    assert_eq!(
        run_prints(
            r#"<?php echo in_array(PHP_OS_FAMILY, ['Linux','Darwin','Windows','FreeBSD','OpenBSD','Solaris','Unknown']) ? 'known' : 'unknown'; "#
        ),
        vec!["known"]
    );
}
#[test]
fn php_eol_constant() {
    assert_eq!(
        run_prints(r#"<?php echo strlen(PHP_EOL) >= 1 ? 'has_eol' : 'no'; "#),
        vec!["has_eol"]
    );
}
#[test]
fn php_maxpathlen() {
    assert_eq!(
        run_prints(r#"<?php echo PHP_MAXPATHLEN > 0 ? 'positive' : 'zero'; "#),
        vec!["positive"]
    );
}

// ── Mathematical constants ────────────────────────────────────

#[test]
fn m_pi_value() {
    assert_eq!(run_prints(r#"<?php echo round(M_PI, 2); "#), vec!["3.14"]);
}
#[test]
fn m_sqrt2_value() {
    assert_eq!(
        run_prints(r#"<?php echo round(M_SQRT2 * M_SQRT2, 5); "#),
        vec!["2"]
    );
}
#[test]
fn m_log2e() {
    assert_eq!(
        run_prints(r#"<?php echo round(M_LOG2E * M_LN2, 5); "#),
        vec!["1"]
    );
}
#[test]
fn nan_inf_constants() {
    assert_eq!(
        run_prints(
            r#"<?php echo is_nan(NAN) ? 'nan' : 'no'; echo is_infinite(INF) ? 'inf' : 'no'; "#
        ),
        vec!["naninf"]
    );
}

// ── Integer / float boundaries ────────────────────────────────

#[test]
fn php_int_max_value() {
    assert_eq!(
        run_prints(r#"<?php echo PHP_INT_MAX === 9223372036854775807 ? 'correct' : 'wrong'; "#),
        vec!["correct"]
    );
}
#[test]
fn php_int_min_value() {
    assert_eq!(
        run_prints(r#"<?php echo PHP_INT_MIN === -9223372036854775808 ? 'correct' : 'wrong'; "#),
        vec!["correct"]
    );
}
#[test]
fn php_float_epsilon_positive() {
    assert_eq!(
        run_prints(
            r#"<?php echo PHP_FLOAT_EPSILON > 0 && PHP_FLOAT_EPSILON < 1 ? 'ok' : 'fail'; "#
        ),
        vec!["ok"]
    );
}

// ── User-defined constants ────────────────────────────────────

#[test]
fn define_const_case_insensitive_default() {
    assert_eq!(
        run_prints(r#"<?php define('MY_CONST', 42); echo MY_CONST; "#),
        vec!["42"]
    );
}
#[test]
fn const_in_class() {
    assert_eq!(
        run_prints(
            r#"<?php class App { const string VERSION = '1.0.0'; const int MAX_RETRY = 3; } echo App::VERSION . ':' . App::MAX_RETRY; "#
        ),
        vec!["1.0.0:3"]
    );
}
#[test]
fn const_in_interface_inherited() {
    assert_eq!(
        run_prints(
            r#"<?php
interface HasMax { const int MAX = 100; }
interface HasMin extends HasMax { const int MIN = 0; }
class Range implements HasMin {}
echo Range::MAX . '-' . Range::MIN;
"#
        ),
        vec!["100-0"]
    );
}

// ── STDIN / STDOUT / STDERR constants ────────────────────────

#[test]
fn php_stdin_defined() {
    assert_eq!(
        run_prints(r#"<?php echo defined('STDIN') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn php_stdout_defined() {
    assert_eq!(
        run_prints(r#"<?php echo defined('STDOUT') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}

// ── Sort flags ────────────────────────────────────────────────

#[test]
fn sort_flags_defined() {
    assert_eq!(
        run_prints(
            r#"<?php
echo defined('SORT_REGULAR') ? '1' : '0';
echo defined('SORT_NUMERIC') ? '1' : '0';
echo defined('SORT_STRING') ? '1' : '0';
echo defined('SORT_NATURAL') ? '1' : '0';
"#
        ),
        vec!["1111"]
    );
}

// ── Array flags ───────────────────────────────────────────────

#[test]
fn array_filter_use_flags_defined() {
    assert_eq!(
        run_prints(
            r#"<?php
echo defined('ARRAY_FILTER_USE_KEY') ? '1' : '0';
echo defined('ARRAY_FILTER_USE_BOTH') ? '1' : '0';
"#
        ),
        vec!["1", "1"]
    );
}

// ── String constants ──────────────────────────────────────────

#[test]
fn directory_separator() {
    assert_eq!(
        run_prints(r#"<?php echo in_array(DIRECTORY_SEPARATOR, ['/', '\\']) ? 'ok' : 'fail'; "#),
        vec!["ok"]
    );
}
#[test]
fn path_separator() {
    assert_eq!(
        run_prints(r#"<?php echo in_array(PATH_SEPARATOR, [':', ';']) ? 'ok' : 'fail'; "#),
        vec!["ok"]
    );
}

// ── JSON constants ────────────────────────────────────────────

#[test]
fn json_constants_defined() {
    assert_eq!(
        run_prints(
            r#"<?php
echo defined('JSON_PRETTY_PRINT') ? '1' : '0';
echo defined('JSON_UNESCAPED_UNICODE') ? '1' : '0';
echo defined('JSON_THROW_ON_ERROR') ? '1' : '0';
"#
        ),
        vec!["111"]
    );
}
