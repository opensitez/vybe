macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_proc_binding_eval,
    "p = proc { a = 1 }; eval('a = 2', p.binding); puts eval('a', p.binding)",
    "2"
);
ruby_test!(
    test_proc_binding_local_variables,
    "a = 1; p = proc { b = 2 }; puts p.binding.local_variables.include?(:a).to_s",
    "true"
);
ruby_test!(
    test_proc_source_location,
    "p = proc {}; puts p.source_location.class.name",
    "Array"
);
ruby_test!(
    test_proc_parameters,
    "p = proc { |a, b=1, *c| }; puts p.parameters.length",
    "3"
);
ruby_test!(
    test_proc_hash,
    "p = proc {}; puts p.hash.class.name",
    "Integer"
);
ruby_test!(test_proc_eql, "p = proc {}; puts p.eql?(p)", "true");
ruby_test!(
    test_proc_eqq,
    "p = proc { |x| x % 2 == 0 }; puts p === 4",
    "true"
);
ruby_test!(
    test_proc_to_s,
    "p = proc {}; puts p.to_s.start_with?('#<Proc:')",
    "true"
);
