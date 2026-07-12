macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_clear_basic,
    "h = {a: 1, b: 2}; h.clear; puts h.length",
    "0"
);
ruby_test!(
    test_clear_returns_self,
    "h = {a: 1}; puts h.clear.object_id == h.object_id",
    "true"
);
ruby_test!(test_clear_empty, "h = {}; h.clear; puts h.length", "0");
ruby_test!(
    test_replace_basic,
    "h = {a: 1}; h.replace({b: 2, c: 3}); puts h.keys.map(&:to_s).join('-')",
    "b-c"
);
ruby_test!(
    test_replace_returns_self,
    "h = {a: 1}; puts h.replace({b: 2}).object_id == h.object_id",
    "true"
);
ruby_test!(
    test_replace_changes_length,
    "h = {a: 1}; h.replace({b: 2, c: 3}); puts h.length",
    "2"
);
ruby_test!(
    test_replace_with_empty,
    "h = {a: 1}; h.replace({}); puts h.length",
    "0"
);
ruby_test!(
    test_replace_self,
    "h = {a: 1}; h.replace(h); puts h.keys.map(&:to_s).join('-')",
    "a"
);
ruby_test!(
    test_replace_preserves_default,
    "h = Hash.new('def'); h.replace({a: 1}); puts h[:b]",
    "def"
);
ruby_test!(
    test_replace_preserves_default_proc,
    "h = Hash.new {|h, k| 'def'}; h.replace({a: 1}); puts h[:b]",
    "def"
);
ruby_test!(
    test_replace_frozen_error,
    "# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.replace({b: 2}); rescue FrozenError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_clear_frozen_error,
    "# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.clear; rescue FrozenError; puts 'err'; end",
    "err"
);
