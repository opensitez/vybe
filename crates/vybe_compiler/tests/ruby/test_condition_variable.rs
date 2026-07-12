macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_condition_variable_basic,
    "m = Mutex.new; cv = ConditionVariable.new; a = 0; t = Thread.new { m.synchronize { a = 1; cv.signal } }; m.synchronize { cv.wait(m) while a == 0 }; puts a",
    "1"
);
ruby_test!(
    test_condition_variable_broadcast,
    "m = Mutex.new; cv = ConditionVariable.new; a = 0; t1 = Thread.new { m.synchronize { cv.wait(m) until a == 1 } }; t2 = Thread.new { m.synchronize { cv.wait(m) until a == 1 } }; m.synchronize { a = 1; cv.broadcast }; t1.join; t2.join; puts a",
    "1"
);
ruby_test!(
    test_condition_variable_wait_timeout,
    "m = Mutex.new; cv = ConditionVariable.new; m.synchronize { cv.wait(m, 0.01) }; puts 'done'",
    "done"
);
