use super::helpers::run_prints;

fn assert_int(expr: &str, expected: i64) {
    assert_eq!(
        run_prints(&format!("<?php echo {}; ", expr)),
        vec![expected.to_string()]
    );
}

#[test]
fn php_spl_structures() {
    for i in 1..=10_i64 {
        assert_int(
            &format!(
                "$q = new SplQueue();\n$q->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO);\n$q->enqueue(10);\n$q->enqueue({i});\n$front = $q->dequeue();\necho $q->top();",
            ),
            i,
        );

        assert_int(
            &format!(
                "$stack = new SplStack();\n$stack->push({i});\n$stack->push({i} + 1);\n$stack->pop();\necho $stack->top();",
            ),
            i,
        );

        assert_int(
            &format!(
                "$p = new SplPriorityQueue();\n$p->setExtractFlags(SplPriorityQueue::EXTR_DATA);\n$p->insert(100 + {i}, {i});\n$p->insert(10, {i} - 1);\necho $p->extract();"
            ),
            100 + i,
        );
    }

    for i in 1..=10_i64 {
        let limit = i + 5;
        assert_int(
            &format!(
                "$objStore = new SplObjectStorage();\n$a = new DateTimeImmutable();\n$b = new DateTimeImmutable();\n$objStore->attach($a, {i});\n$objStore->attach($b, {i} + {limit});\necho $objStore[$a] + $objStore[$b];"
            ),
            2 * i + 5,
        );

        assert_int(
            &format!(
                "$list = new SplFixedArray({limit});\nfor ($idx = 0; $idx < {limit}; $idx++) {{ $list[$idx] = $idx; }}\necho $list[{i}] + $list[{i} - 1];",
                limit = limit,
                i = i
            ),
            2 * i - 1,
        );
    }
}
