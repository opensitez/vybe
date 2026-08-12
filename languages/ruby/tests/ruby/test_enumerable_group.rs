macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_enumerable_group_by_basic,
    "puts [1, 2, 3, 4].group_by { |x| x % 2 }[0].join('-')",
    "2-4"
);
ruby_test!(
    test_enumerable_group_by_keys,
    "puts [1, 2, 3, 4].group_by { |x| x % 2 }.keys.sort.join('-')",
    "0-1"
);
ruby_test!(
    test_enumerable_partition_basic,
    "puts [1, 2, 3, 4].partition { |x| x.even? }.map { |a| a.join('-') }.join('|')",
    "2-4|1-3"
);
ruby_test!(
    test_enumerable_slice_before_regex,
    "puts ['a', '1', 'b', '2'].slice_before(/[0-9]/).map{|a| a.join('-')}.join('|')",
    "a|1-b|2"
);
ruby_test!(
    test_enumerable_slice_before_block,
    "puts [1, 2, 3, 4].slice_before { |x| x.even? }.map{|a| a.join('-')}.join('|')",
    "1|2-3|4"
);
ruby_test!(
    test_enumerable_slice_after_regex,
    "puts ['a', '1', 'b', '2'].slice_after(/[0-9]/).map{|a| a.join('-')}.join('|')",
    "a-1|b-2"
);
ruby_test!(
    test_enumerable_slice_after_block,
    "puts [1, 2, 3, 4].slice_after { |x| x.even? }.map{|a| a.join('-')}.join('|')",
    "1-2|3-4"
);
ruby_test!(
    test_enumerable_slice_when_basic,
    "puts [1, 2, 4, 5, 8].slice_when { |i, j| i+1 != j }.map{|a| a.join('-')}.join('|')",
    "1-2|4-5|8"
);
ruby_test!(
    test_enumerable_chunk_basic,
    "puts [1, 2, 2, 3].chunk { |x| x.even? }.map{|k, v| \"#{k}:#{v.join('-')}\"}.join('|')",
    "false:1|true:2-2|false:3"
);
ruby_test!(
    test_enumerable_chunk_drop,
    "puts [1, 2, 2, 3].chunk { |x| x.even? ? true : nil }.map{|k, v| \"#{k}:#{v.join('-')}\"}.join('|')",
    "true:2-2"
); // nil drops the chunk
