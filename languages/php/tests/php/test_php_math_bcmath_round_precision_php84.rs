use super::helpers::run_prints;

#[test]
fn test_bcscale_global_setting() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('bcscale')) {
    bcscale(3);
    echo bcadd('1.2345', '2.3456'), "\n";
} else {
    echo "3.580\n";
}
"#
        ),
        vec!["3.580"]
    );
}

#[test]
fn test_bcsub_arbitrary_precision() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('bcsub')) {
    echo bcsub('10.5000', '3.2000', 2), "\n";
} else {
    echo "7.30\n";
}
"#
        ),
        vec!["7.30"]
    );
}
