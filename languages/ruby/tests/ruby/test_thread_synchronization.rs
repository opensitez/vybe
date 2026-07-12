macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_thread_synchronization_mutex,
    "m = Mutex.new; a = 0; t1 = Thread.new { m.synchronize { a += 1 } }; t2 = Thread.new { m.synchronize { a += 1 } }; t1.join; t2.join; puts a",
    "2"
);
ruby_test!(
    test_thread_synchronization_mutex_locked,
    "m = Mutex.new; m.lock; puts m.locked?; m.unlock",
    "true"
);
ruby_test!(
    test_thread_synchronization_mutex_try_lock,
    "m = Mutex.new; puts m.try_lock; m.unlock",
    "true"
);
ruby_test!(
    test_thread_synchronization_mutex_owned,
    "m = Mutex.new; m.lock; puts m.owned?; m.unlock",
    "true"
);
ruby_test!(
    test_thread_group_list,
    "puts ThreadGroup::Default.list.include?(Thread.current).to_s",
    "true"
);
ruby_test!(
    test_thread_group_add,
    "g = ThreadGroup.new; t = Thread.new { sleep 0.1 }; g.add(t); puts g.list.include?(t).to_s",
    "true"
);
ruby_test!(
    test_thread_group_enclose,
    "g = ThreadGroup.new; g.enclose; puts g.enclosed?",
    "true"
);
