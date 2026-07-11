
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_thread_creation, "t = Thread.new { 42 }; puts t.value", "42");
ruby_test!(test_thread_join, "t = Thread.new { sleep 0.1; 42 }; puts t.join.class.name", "Thread");
ruby_test!(test_thread_status_run, "t = Thread.new { sleep 1 }; puts t.status", "run"); // or sleep
ruby_test!(test_thread_status_false, "t = Thread.new { 42 }; t.join; puts t.status.inspect", "false");
ruby_test!(test_thread_alive, "t = Thread.new { sleep 1 }; puts t.alive?", "true");
ruby_test!(test_thread_alive_false, "t = Thread.new { 42 }; t.join; puts t.alive?", "false");
ruby_test!(test_thread_current, "puts Thread.current.class.name", "Thread");
ruby_test!(test_thread_main, "puts Thread.main.class.name", "Thread");
ruby_test!(test_thread_list, "puts Thread.list.class.name", "Array");
ruby_test!(test_thread_variables, "Thread.current[:my_var] = 42; puts Thread.current[:my_var]", "42");
ruby_test!(test_thread_key_check, "Thread.current[:my_var] = 42; puts Thread.current.key?(:my_var)", "true");
ruby_test!(test_thread_keys, "Thread.current[:my_var] = 42; puts Thread.current.keys.include?(:my_var)", "true");
ruby_test!(test_thread_kill, "t = Thread.new { sleep 10 }; t.kill; t.join; puts t.alive?", "false");
