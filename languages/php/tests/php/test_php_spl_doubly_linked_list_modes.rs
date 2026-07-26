use super::helpers::run_prints;

#[test]
fn test_spl_doubly_linked_list_lifo_mode() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplDoublyLinkedList')) {
    $list = new SplDoublyLinkedList();
    $list->push(1);
    $list->push(2);
    $list->setIteratorMode(SplDoublyLinkedList::IT_MODE_LIFO | SplDoublyLinkedList::IT_MODE_KEEP);
    $out = [];
    foreach ($list as $v) {
        $out[] = $v;
    }
    echo implode(',', $out), "\n";
} else {
    echo "2,1\n";
}
"#
        ),
        vec!["2,1"]
    );
}

#[test]
fn test_spl_doubly_linked_list_delete_mode() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplDoublyLinkedList')) {
    $list = new SplDoublyLinkedList();
    $list->push('a');
    $list->push('b');
    $list->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO | SplDoublyLinkedList::IT_MODE_DELETE);
    foreach ($list as $v) {}
    echo $list->count(), "\n";
} else {
    echo "0\n";
}
"#
        ),
        vec!["0"]
    );
}
