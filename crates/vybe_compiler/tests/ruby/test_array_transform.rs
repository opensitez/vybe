use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_array_transform_map, "puts [1, 2, 3].map { |x| x * 2 }.join('-')", "2-4-6");
ruby_test!(test_array_transform_collect, "puts [1, 2, 3].collect { |x| x * 2 }.join('-')", "2-4-6");
ruby_test!(test_array_transform_map_bang, "a = [1, 2, 3]; a.map! { |x| x * 2 }; puts a.join('-')", "2-4-6");
ruby_test!(test_array_transform_filter, "puts [1, 2, 3, 4].filter { |x| x.even? }.join('-')", "2-4");
ruby_test!(test_array_transform_select, "puts [1, 2, 3, 4].select { |x| x.even? }.join('-')", "2-4");
ruby_test!(test_array_transform_select_bang, "a = [1, 2, 3, 4]; a.select! { |x| x.even? }; puts a.join('-')", "2-4");
ruby_test!(test_array_transform_reject, "puts [1, 2, 3, 4].reject { |x| x.even? }.join('-')", "1-3");
ruby_test!(test_array_transform_reject_bang, "a = [1, 2, 3, 4]; a.reject! { |x| x.even? }; puts a.join('-')", "1-3");
ruby_test!(test_array_transform_filter_map, "puts [1, 2, 3, 4].filter_map { |x| x * 2 if x.even? }.join('-')", "4-8");
ruby_test!(test_array_transform_compact, "puts [1, nil, 3].compact.join('-')", "1-3");
ruby_test!(test_array_transform_compact_bang, "a = [1, nil, 3]; a.compact!; puts a.join('-')", "1-3");
ruby_test!(test_array_transform_map_enumerator, "puts [1].map.class.name", "Enumerator");
ruby_test!(test_array_transform_select_enumerator, "puts [1].select.class.name", "Enumerator");
