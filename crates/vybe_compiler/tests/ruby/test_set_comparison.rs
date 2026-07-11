
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_set_subset, "require 'set'; s1 = Set.new([1]); s2 = Set.new([1, 2]); puts s1.subset?(s2)", "true");
ruby_test!(test_set_proper_subset, "require 'set'; s1 = Set.new([1]); s2 = Set.new([1]); puts s1.proper_subset?(s2)", "false");
ruby_test!(test_set_superset, "require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([1]); puts s1.superset?(s2)", "true");
ruby_test!(test_set_proper_superset, "require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([1, 2]); puts s1.proper_superset?(s2)", "false");
ruby_test!(test_set_equality, "require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 1]); puts s1 == s2", "true");
