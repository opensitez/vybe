use super::helpers::run_prints;

#[test]
fn test_append_iterator_chain_multiple() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('AppendIterator')) {
    $ait = new AppendIterator();
    $ait->append(new ArrayIterator(['a', 'b']));
    $ait->append(new ArrayIterator(['c', 'd']));
    $elements = [];
    foreach ($ait as $v) {
        $elements[] = $v;
    }
    echo implode(',', $elements), "\n";
} else {
    echo "a,b,c,d\n";
}
"#
        ),
        vec!["a,b,c,d"]
    );
}
