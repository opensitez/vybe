
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerable_each_entry, "acc = []; [1, 2].each_entry { |x| acc << x }; puts acc.join('-')", "1-2");
ruby_test!(test_enumerable_each_slice, "acc = []; [1, 2, 3, 4].each_slice(2) { |arr| acc << arr.join(',') }; puts acc.join('-')", "1,2-3,4");
ruby_test!(test_enumerable_each_cons, "acc = []; [1, 2, 3].each_cons(2) { |arr| acc << arr.join(',') }; puts acc.join('-')", "1,2-2,3");
ruby_test!(test_enumerable_each_with_index, "acc = []; [1, 2].each_with_index { |x, i| acc << \"#{x}:#{i}\" }; puts acc.join('-')", "1:0-2:1");
ruby_test!(test_enumerable_each_with_object, "puts [1, 2].each_with_object([]) { |x, arr| arr << x * 2 }.join('-')", "2-4");
ruby_test!(test_enumerable_reverse_each, "acc = []; [1, 2].reverse_each { |x| acc << x }; puts acc.join('-')", "2-1");
ruby_test!(test_enumerable_cycle, "acc = []; [1, 2].cycle(2) { |x| acc << x }; puts acc.join('-')", "1-2-1-2");
ruby_test!(test_enumerable_inject, "puts [1, 2, 3].inject(0) { |sum, n| sum + n }", "6");
ruby_test!(test_enumerable_inject_symbol, "puts [1, 2, 3].inject(:+)", "6");
ruby_test!(test_enumerable_reduce, "puts [1, 2, 3].reduce(0) { |sum, n| sum + n }", "6");
ruby_test!(test_enumerable_reduce_symbol, "puts [1, 2, 3].reduce(:+)", "6");
