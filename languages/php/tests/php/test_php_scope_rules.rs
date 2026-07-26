use super::helpers::run_prints;

fn assert_int(expr: &str, expected: i64) {
    assert_eq!(run_prints(&format!("<?php echo {}; ", expr)), vec![expected.to_string()]);
}

#[test]
fn php_scope_rules() {
    for n in 1..=20_i64 {
        let n1 = n;
        let n2 = n + 10;

        assert_int(
            &format!(
                "$seed = {n1};\n$fn = fn() => $seed;\n$seed = 0;\necho $fn();"
            ),
            n1,
        );

        assert_int(
            &format!(
                "$seed = {n2};\n$fn = function() use (&$seed) {{ return $seed; }};\n$seed += 1;\necho $fn();"
            ),
            n2 + 1,
        );

        assert_int(
            &format!(
                "function scoped_static_{n1}() {{ static $count = 0; $count += {n1}; return $count; }}\necho scoped_static_{n1}() + scoped_static_{n1}();"
            ),
            n1 * 2,
        );
    }
}
