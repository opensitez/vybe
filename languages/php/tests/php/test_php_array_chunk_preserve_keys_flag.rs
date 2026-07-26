use super::helpers::run_prints;

#[test]
fn test_array_chunk_preserve_numeric_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$input = [10 => 'a', 20 => 'b', 30 => 'c'];
$chunks = array_chunk($input, 2, true);
echo implode(',', array_keys($chunks[0])) . '|' . implode(',', array_keys($chunks[1])), "\n";
"#
        ),
        vec!["10,20|30"]
    );
}

#[test]
fn test_array_chunk_reindex_keys_default() {
    assert_eq!(
        run_prints(
            r#"<?php
$input = [10 => 'a', 20 => 'b', 30 => 'c'];
$chunks = array_chunk($input, 2, false);
echo implode(',', array_keys($chunks[0])) . '|' . implode(',', array_keys($chunks[1])), "\n";
"#
        ),
        vec!["0,1|0"]
    );
}
