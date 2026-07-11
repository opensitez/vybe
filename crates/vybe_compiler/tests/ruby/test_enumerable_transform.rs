
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerable_map, "puts [1, 2, 3].map { |x| x * 2 }.join('-')", "2-4-6");
ruby_test!(test_enumerable_collect, "puts [1, 2, 3].collect { |x| x * 2 }.join('-')", "2-4-6");
ruby_test!(test_enumerable_collect_concat, "puts [1, 2].collect_concat { |x| [x, x] }.join('-')", "1-1-2-2");
ruby_test!(test_enumerable_flat_map, "puts [1, 2].flat_map { |x| [x, x] }.join('-')", "1-1-2-2");
ruby_test!(test_enumerable_filter, "puts [1, 2, 3, 4].filter { |x| x.even? }.join('-')", "2-4");
ruby_test!(test_enumerable_select, "puts [1, 2, 3, 4].select { |x| x.even? }.join('-')", "2-4");
ruby_test!(test_enumerable_reject, "puts [1, 2, 3, 4].reject { |x| x.even? }.join('-')", "1-3");
ruby_test!(test_enumerable_filter_map, "puts [1, 2, 3, 4].filter_map { |x| x * 2 if x.even? }.join('-')", "4-8");
ruby_test!(test_enumerable_partition, "puts [1, 2, 3, 4].partition { |x| x.even? }.map { |arr| arr.join(',') }.join('-')", "2,4-1,3");
ruby_test!(test_enumerable_zip, "puts [1, 2].zip(['a', 'b']).map { |arr| arr.join(',') }.join('-')", "1,a-2,b");
ruby_test!(test_enumerable_zip_block, "acc = []; [1, 2].zip(['a', 'b']) { |x, y| acc << \"#{x}:#{y}\" }; puts acc.join('-')", "1:a-2:b");
