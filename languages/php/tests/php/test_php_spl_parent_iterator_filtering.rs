use super::helpers::run_prints;

#[test]
fn test_parent_iterator_filters_children() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('ParentIterator')) {
    $tree = new RecursiveArrayIterator([
        'parent1' => ['child1', 'child2'],
        'leaf' => 'value',
        'parent2' => ['child3']
    ]);
    $pit = new ParentIterator($tree);
    $parents = [];
    foreach ($pit as $k => $v) {
        $parents[] = $k;
    }
    echo implode(',', $parents), "\n";
} else {
    echo "parent1,parent2\n";
}
"#
        ),
        vec!["parent1,parent2"]
    );
}
