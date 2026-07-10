use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_thread_basic, "t = Thread.new { 1 + 2 }; puts t.value", "3");
ruby_test!(test_thread_join, "t = Thread.new { sleep 0.01; 'done' }; t.join; puts t.value", "done");
ruby_test!(test_thread_current, "puts Thread.current.class.name", "Thread");
ruby_test!(test_thread_main, "puts Thread.main == Thread.current", "true");
ruby_test!(test_thread_status, "t = Thread.new { sleep 0.01 }; puts t.status.is_a?(String)", "true");
ruby_test!(test_thread_alive, "t = Thread.new { sleep 0.01 }; puts t.alive?", "true");
ruby_test!(test_thread_pass, "Thread.pass; puts 'ok'", "ok");
ruby_test!(test_thread_variables, "t = Thread.current; t[:my_var] = 123; puts t[:my_var]", "123");
ruby_test!(test_thread_key_check, "t = Thread.current; t[:my_var] = 123; puts t.key?(:my_var)", "true");
ruby_test!(test_thread_keys, "t = Thread.current; t[:my_var] = 123; puts t.keys.include?(:my_var)", "true");
