
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_sized_queue_creation, "q = SizedQueue.new(2); puts q.max", "2");
ruby_test!(test_sized_queue_max_set, "q = SizedQueue.new(2); q.max = 5; puts q.max", "5");
ruby_test!(test_sized_queue_push_pop, "q = SizedQueue.new(2); q.push(1); puts q.pop", "1");
ruby_test!(test_sized_queue_empty, "q = SizedQueue.new(2); puts q.empty?", "true");
ruby_test!(test_sized_queue_clear, "q = SizedQueue.new(2); q.push(1); q.clear; puts q.empty?", "true");
ruby_test!(test_sized_queue_length, "q = SizedQueue.new(2); q.push(1); q.push(2); puts q.length", "2");
ruby_test!(test_sized_queue_num_waiting, "q = SizedQueue.new(2); puts q.num_waiting", "0");
ruby_test!(test_sized_queue_close, "q = SizedQueue.new(2); q.close; puts q.closed?", "true");
ruby_test!(test_sized_queue_push_closed, "q = SizedQueue.new(2); q.close; begin; q.push(1); rescue ClosedQueueError; puts 'err'; end", "err");
ruby_test!(test_sized_queue_pop_non_block, "q = SizedQueue.new(2); begin; q.pop(true); rescue ThreadError; puts 'err'; end", "err");
ruby_test!(test_sized_queue_push_non_block, "q = SizedQueue.new(1); q.push(1); begin; q.push(2, true); rescue ThreadError; puts 'err'; end", "err");
