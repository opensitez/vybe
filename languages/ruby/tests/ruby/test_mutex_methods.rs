macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_mutex_new, "m = Mutex.new; puts m.class.name", "Mutex");
ruby_test!(
    test_mutex_lock,
    "m = Mutex.new; m.lock; puts m.locked?",
    "true"
);
ruby_test!(
    test_mutex_unlock,
    "m = Mutex.new; m.lock; m.unlock; puts m.locked?",
    "false"
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
    test_mutex_synchronize,
    "m = Mutex.new; puts m.synchronize { 42 }",
    "42"
);
ruby_test!(
    test_mutex_owned,
    "m = Mutex.new; m.lock; puts m.owned?",
    "true"
);
ruby_test!(
    test_mutex_owned_false,
    "m = Mutex.new; puts m.owned?",
    "false"
);
ruby_test!(
    test_mutex_sleep,
    "m = Mutex.new; m.lock; puts m.sleep(0.01).class.name",
    "Integer"
);
