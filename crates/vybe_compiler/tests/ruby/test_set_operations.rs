
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_set_union, "require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 3]); puts (s1 | s2).to_a.sort.join('-')", "1-2-3");
ruby_test!(test_set_intersection, "require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 3]); puts (s1 & s2).to_a.sort.join('-')", "2");
ruby_test!(test_set_difference, "require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 3]); puts (s1 - s2).to_a.sort.join('-')", "1");
ruby_test!(test_set_xor, "require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 3]); puts (s1 ^ s2).to_a.sort.join('-')", "1-3");
ruby_test!(test_set_disjoint, "require 'set'; s1 = Set.new([1]); s2 = Set.new([2]); puts s1.disjoint?(s2)", "true");
ruby_test!(test_set_intersect, "require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 3]); puts s1.intersect?(s2)", "true");
