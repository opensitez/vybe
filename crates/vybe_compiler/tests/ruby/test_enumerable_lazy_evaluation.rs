
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_lazy_chaining, "puts [1, 2, 3, 4].lazy.select {|x| x % 2 == 0}.map {|x| x * 10}.force.join('-')", "20-40");
ruby_test!(test_lazy_infinite_sequence, "puts (1..Float::INFINITY).lazy.select {|x| x % 2 == 0}.first(3).join('-')", "2-4-6");
ruby_test!(test_lazy_enum_for, "puts [1, 2].enum_for(:each).lazy.map {|x| x * 2}.force.join('-')", "2-4");
ruby_test!(test_lazy_reject_map, "puts [1, 2, 3].lazy.reject {|x| x == 2}.map {|x| x * 2}.force.join('-')", "2-6");
ruby_test!(test_lazy_grep_map, "puts [1, 'a', 2].lazy.grep(Integer).map {|x| x * 2}.force.join('-')", "2-4");
ruby_test!(test_lazy_take, "puts (1..Float::INFINITY).lazy.take(3).force.join('-')", "1-2-3");
ruby_test!(test_lazy_drop, "puts (1..5).lazy.drop(3).force.join('-')", "4-5");
ruby_test!(test_lazy_flat_map_chain, "puts [1, 2].lazy.flat_map {|x| [x, x]}.map {|x| x * 10}.force.join('-')", "10-10-20-20");
ruby_test!(test_lazy_zip_chain, "puts [1, 2].lazy.zip([3, 4]).map {|x, y| x + y}.force.join('-')", "4-6");
