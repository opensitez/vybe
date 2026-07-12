macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_module_ancestors_basic,
    "module M; end; class C; include M; end; puts C.ancestors.include?(M).to_s",
    "true"
);
ruby_test!(
    test_module_ancestors_prepend,
    "module M; end; class C; prepend M; end; puts C.ancestors.first == M",
    "true"
);
ruby_test!(
    test_module_included_modules,
    "module M; end; class C; include M; end; puts C.included_modules.include?(M).to_s",
    "true"
);
ruby_test!(
    test_module_include_question,
    "module M; end; class C; include M; end; puts C.include?(M)",
    "true"
);
ruby_test!(
    test_module_include_question_false,
    "module M; end; class C; end; puts C.include?(M)",
    "false"
);
ruby_test!(
    test_module_name,
    "module MyMod; end; puts MyMod.name",
    "MyMod"
);
ruby_test!(
    test_module_name_anonymous,
    "puts Module.new.name.nil?",
    "true"
);
ruby_test!(
    test_module_to_s,
    "module MyMod; end; puts MyMod.to_s",
    "MyMod"
);
ruby_test!(
    test_module_inspect,
    "module MyMod; end; puts MyMod.inspect",
    "MyMod"
);
