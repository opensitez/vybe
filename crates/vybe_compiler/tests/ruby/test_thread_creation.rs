
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_thread_creation_basic, "t = Thread.new { 1 + 1 }; puts t.value", "2");
ruby_test!(test_thread_creation_args, "t = Thread.new(10) { |x| x * 2 }; puts t.value", "20");
ruby_test!(test_thread_join, "t = Thread.new { sleep 0.1; 42 }; puts t.join.value", "42");
ruby_test!(test_thread_status, "t = Thread.new { sleep 0.1 }; puts %w[run sleep].include?(t.status).to_s", "true");
ruby_test!(test_thread_alive, "t = Thread.new { sleep 0.1 }; puts t.alive?", "true");
ruby_test!(test_thread_main, "puts Thread.main == Thread.current", "true");
ruby_test!(test_thread_current, "puts Thread.current.class.name", "Thread");
ruby_test!(test_thread_list, "puts Thread.list.include?(Thread.current).to_s", "true");
ruby_test!(test_thread_name, "t = Thread.new {}; t.name = 'worker'; puts t.name", "worker");
ruby_test!(test_thread_keys, "t = Thread.current; t[:foo] = 'bar'; puts t.keys.include?(:foo).to_s", "true");
ruby_test!(test_thread_key_question, "t = Thread.current; t[:baz] = 1; puts t.key?(:baz)", "true");
ruby_test!(test_thread_fetch, "t = Thread.current; t[:qux] = 42; puts t.fetch(:qux)", "42");
