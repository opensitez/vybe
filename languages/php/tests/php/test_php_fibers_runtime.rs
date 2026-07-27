use super::helpers::run_prints;

fn assert_output(expr: &str, expected: &str) {
    assert_eq!(
        run_prints(&format!("<?php echo {}; ", expr)),
        vec![expected.to_string()]
    );
}

fn assert_int(expr: &str, expected: i64) {
    assert_output(expr, &expected.to_string());
}

#[test]
fn php_fiber_async_runtime() {
    for n in 1..=20_i64 {
        let total = n * (n - 1) / 2;
        let resumed = n * 2;
        assert_int(
            &format!(
                "$f = new Fiber(function(int $n) {{\n    $sum = 0;\n    for ($i = 0; $i < $n; $i++) {{\n        $sum += $i;\n    }}\n    return $sum;\n}});\n$f->start({n});\necho $f->getReturn();"
            ),
            total,
        );
        assert_int(
            &format!(
                "$f = new Fiber(function(int $n) {{\n    return $n * 2;\n}});\n$f->start({n});\necho $f->getReturn();"
            ),
            resumed,
        );
    }
}
