macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_queue_push, "q = Queue.new; q.push(1); puts q.pop", "1");
ruby_test!(test_queue_enq, "q = Queue.new; q.enq(1); puts q.deq", "1");
ruby_test!(test_queue_empty, "q = Queue.new; puts q.empty?", "true");
ruby_test!(
    test_queue_clear,
    "q = Queue.new; q.push(1); q.clear; puts q.empty?",
    "true"
);
ruby_test!(
    test_queue_length,
    "q = Queue.new; q.push(1); q.push(2); puts q.length",
    "2"
);
ruby_test!(
    test_queue_size,
    "q = Queue.new; q.push(1); q.push(2); puts q.size",
    "2"
);
ruby_test!(
    test_queue_num_waiting,
    "q = Queue.new; puts q.num_waiting",
    "0"
);
ruby_test!(
    test_queue_close,
    "q = Queue.new; q.close; puts q.closed?",
    "true"
);
ruby_test!(
    test_queue_push_closed,
    "q = Queue.new; q.close; begin; q.push(1); rescue ClosedQueueError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_queue_pop_non_block,
    "q = Queue.new; begin; q.pop(true); rescue ThreadError; puts 'err'; end",
    "err"
);
