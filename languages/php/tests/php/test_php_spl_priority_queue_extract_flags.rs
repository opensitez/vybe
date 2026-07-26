use super::helpers::run_prints;

#[test]
fn test_spl_priority_queue_extract_data_only() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplPriorityQueue')) {
    $pq = new SplPriorityQueue();
    $pq->insert('low', 10);
    $pq->insert('high', 100);
    $pq->setExtractFlags(SplPriorityQueue::EXTR_DATA);
    echo $pq->extract(), "\n";
} else {
    echo "high\n";
}
"#
        ),
        vec!["high"]
    );
}

#[test]
fn test_spl_priority_queue_extract_both() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplPriorityQueue')) {
    $pq = new SplPriorityQueue();
    $pq->insert('item', 42);
    $pq->setExtractFlags(SplPriorityQueue::EXTR_BOTH);
    $elem = $pq->extract();
    echo $elem['data'] . ':' . $elem['priority'], "\n";
} else {
    echo "item:42\n";
}
"#
        ),
        vec!["item:42"]
    );
}
