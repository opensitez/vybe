use super::helpers::run_prints;

fn assert_int(expr: &str, expected: i64) {
    assert_eq!(run_prints(&format!("<?php echo {}; ", expr)), vec![expected.to_string()]);
}

fn int_array(len: i64) -> String {
    let values: Vec<String> = (0..len).map(|v| v.to_string()).collect();
    format!("[{}]", values.join(", "))
}

#[test]
fn php_array_surface_features() {
    for len in 1..=20_i64 {
        let arr = int_array(len);
        let sum = len * (len - 1) / 2;
        let odd_count = (len + 1) / 2;
        let even_count = len / 2;
        assert_int(&format!("count({arr})"), len);
        assert_int(&format!("array_sum({arr})"), sum);
        assert_int(&format!("count(array_filter({arr}, fn($v) => $v % 2 === 0))"), even_count);
        assert_int(&format!("count(array_filter({arr}, fn($v) => $v % 2 === 1))"), odd_count);
        assert_int(&format!("array_key_first({arr})"), 0);
        assert_int(&format!("array_key_last({arr})"), len - 1);
        assert_int(
            &format!("array_reduce({arr}, fn($carry, $item) => $carry + $item, 0)"),
            sum,
        );
    }
}
