macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_mutex_basic,
    "m = Mutex.new; x = 0; m.synchronize { x = 1 }; puts x",
    "1"
);
ruby_test!(
    test_mutex_lock_unlock,
    "m = Mutex.new; m.lock; puts m.locked?; m.unlock; puts m.locked?",
    "true\nfalse"
); // Wait, run_ruby_one just expects last expression.
ruby_test!(
    test_mutex_locked,
    "m = Mutex.new; m.lock; puts m.locked?",
    "true"
);
ruby_test!(
    test_mutex_try_lock,
    "m = Mutex.new; puts m.try_lock",
    "true"
);
ruby_test!(
    test_mutex_try_lock_fail,
    "m = Mutex.new; m.lock; puts m.try_lock",
    "false"
);
ruby_test!(
    test_mutex_owned,
    "m = Mutex.new; m.lock; puts m.owned?",
    "true"
);
