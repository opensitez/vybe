use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_lazy_basic, "puts [1, 2, 3].lazy.map {|x| x * 2}.first(2).join('-')", "2-4");
ruby_test!(test_lazy_class, "puts [1, 2, 3].lazy.class.name", "Enumerator::Lazy");
ruby_test!(test_lazy_does_not_evaluate_all, "acc = []; [1, 2, 3].lazy.map {|x| acc << x; x * 2}.first(2); puts acc.join('-')", "1-2"); // only evaluated up to 2
ruby_test!(test_lazy_select, "puts [1, 2, 3, 4].lazy.select {|x| x % 2 == 0}.first(1).join('-')", "2");
ruby_test!(test_lazy_reject, "puts [1, 2, 3, 4].lazy.reject {|x| x % 2 == 0}.first(1).join('-')", "1");
ruby_test!(test_lazy_grep, "puts [1, 'a', 2].lazy.grep(Integer).first(1).join('-')", "1");
ruby_test!(test_lazy_grep_v, "puts [1, 'a', 2].lazy.grep_v(Integer).first(1).join('-')", "a");
ruby_test!(test_lazy_zip, "puts [1, 2].lazy.zip([3, 4]).first(1).inspect", "[[1, 3]]");
ruby_test!(test_lazy_take_while, "puts [1, 2, 3].lazy.take_while {|x| x < 3}.force.join('-')", "1-2");
ruby_test!(test_lazy_drop_while, "puts [1, 2, 3].lazy.drop_while {|x| x < 3}.force.join('-')", "3");
ruby_test!(test_lazy_flat_map, "puts [1, 2].lazy.flat_map {|x| [x, x]}.force.join('-')", "1-1-2-2");
ruby_test!(test_lazy_force, "puts [1, 2].lazy.map {|x| x * 2}.force.join('-')", "2-4"); // force evaluates fully and returns array
ruby_test!(test_lazy_to_a, "puts [1, 2].lazy.map {|x| x * 2}.to_a.join('-')", "2-4");
