use super::helpers::run_prints;

#[test]
fn test_hrtime_as_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$t = hrtime();
echo is_array($t) && count($t) === 2 ? 'array_ok' : 'err', "\n";
"#
        ),
        vec!["array_ok"]
    );
}

#[test]
fn test_hrtime_as_number() {
    assert_eq!(
        run_prints(
            r#"<?php
$t = hrtime(true);
echo (is_int($t) || is_float($t)) && $t > 0 ? 'number_ok' : 'err', "\n";
"#
        ),
        vec!["number_ok"]
    );
}

#[test]
fn test_hrtime_monotonicity() {
    assert_eq!(
        run_prints(
            r#"<?php
$t1 = hrtime(true);
usleep(100);
$t2 = hrtime(true);
echo $t2 >= $t1 ? 'monotonic' : 'err', "\n";
"#
        ),
        vec!["monotonic"]
    );
}
