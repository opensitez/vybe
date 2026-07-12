macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_threadgroup_add,
    "tg = ThreadGroup.new; t = Thread.new { sleep(0.01) }; tg.add(t); puts tg.list.include?(t).to_s",
    "true"
);
ruby_test!(
    test_threadgroup_enclose,
    "tg = ThreadGroup.new; tg.enclose; puts tg.enclosed?",
    "true"
);
ruby_test!(
    test_threadgroup_add_enclosed,
    "tg = ThreadGroup.new; tg.enclose; t = Thread.new { sleep(0.01) }; begin; tg.add(t); rescue ThreadError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_threadgroup_default,
    "puts ThreadGroup::Default.class.name",
    "ThreadGroup"
);
ruby_test!(
    test_threadgroup_list_default,
    "puts ThreadGroup::Default.list.include?(Thread.main).to_s",
    "true"
);
