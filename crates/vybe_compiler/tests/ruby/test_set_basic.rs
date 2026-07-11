
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_set_creation, "require 'set'; s = Set.new([1, 2, 2, 3]); puts s.to_a.sort.join('-')", "1-2-3");
ruby_test!(test_set_add, "require 'set'; s = Set.new; s.add(1); s.add(1); puts s.size", "1");
ruby_test!(test_set_delete, "require 'set'; s = Set.new([1, 2]); s.delete(1); puts s.to_a.join('-')", "2");
ruby_test!(test_set_include, "require 'set'; s = Set.new([1]); puts s.include?(1)", "true");
ruby_test!(test_set_clear, "require 'set'; s = Set.new([1]); s.clear; puts s.empty?", "true");
ruby_test!(test_set_replace, "require 'set'; s = Set.new([1]); s.replace([2, 3]); puts s.to_a.sort.join('-')", "2-3");
ruby_test!(test_set_merge, "require 'set'; s = Set.new([1]); s.merge([2, 1]); puts s.to_a.sort.join('-')", "1-2");
