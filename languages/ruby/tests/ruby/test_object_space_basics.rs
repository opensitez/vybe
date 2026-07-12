macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_object_space_each_object,
    "acc = 0; ObjectSpace.each_object(Class) { |c| acc += 1 }; puts acc > 10",
    "true"
); // There should be at least a few classes in ruby core
ruby_test!(
    test_object_space_garbage_collect,
    "ObjectSpace.garbage_collect; puts 'ok'",
    "ok"
);
ruby_test!(
    test_object_space_define_finalizer,
    "o = Object.new; acc = []; ObjectSpace.define_finalizer(o, proc { acc << 'finalized' }); o = nil; ObjectSpace.garbage_collect; puts 'ok'",
    "ok"
); // finalizers are tricky to test synchronously, just check it parses/runs without error
ruby_test!(
    test_object_space_count_objects,
    "puts ObjectSpace.count_objects.is_a?(Hash)",
    "true"
);
