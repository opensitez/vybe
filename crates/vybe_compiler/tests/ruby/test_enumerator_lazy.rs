macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_enumerator_lazy_basic,
    "puts [1, 2, 3].lazy.class.name",
    "Enumerator::Lazy"
);
ruby_test!(
    test_enumerator_lazy_map,
    "puts [1, 2, 3].lazy.map { |x| x * 2 }.class.name",
    "Enumerator::Lazy"
);
ruby_test!(
    test_enumerator_lazy_force,
    "puts [1, 2, 3].lazy.map { |x| x * 2 }.force.join('-')",
    "2-4-6"
);
ruby_test!(
    test_enumerator_lazy_select,
    "puts [1, 2, 3, 4].lazy.select { |x| x.even? }.force.join('-')",
    "2-4"
);
ruby_test!(
    test_enumerator_lazy_reject,
    "puts [1, 2, 3, 4].lazy.reject { |x| x.even? }.force.join('-')",
    "1-3"
);
ruby_test!(
    test_enumerator_lazy_grep,
    "puts ['a', 'b', 1].lazy.grep(String).force.join('-')",
    "a-b"
);
ruby_test!(
    test_enumerator_lazy_grep_v,
    "puts ['a', 'b', 1].lazy.grep_v(String).force.join('-')",
    "1"
);
ruby_test!(
    test_enumerator_lazy_take,
    "puts (1..Float::INFINITY).lazy.take(3).force.join('-')",
    "1-2-3"
);
ruby_test!(
    test_enumerator_lazy_take_while,
    "puts [1, 2, 3, 4, 1].lazy.take_while { |x| x < 3 }.force.join('-')",
    "1-2"
);
ruby_test!(
    test_enumerator_lazy_drop,
    "puts [1, 2, 3, 4].lazy.drop(2).force.join('-')",
    "3-4"
);
ruby_test!(
    test_enumerator_lazy_drop_while,
    "puts [1, 2, 3, 4, 1].lazy.drop_while { |x| x < 3 }.force.join('-')",
    "3-4-1"
);
ruby_test!(
    test_enumerator_lazy_flat_map,
    "puts [1, 2].lazy.flat_map { |x| [x, x] }.force.join('-')",
    "1-1-2-2"
);
ruby_test!(
    test_enumerator_lazy_zip,
    "puts [1, 2].lazy.zip(['a', 'b']).force.map { |arr| arr.join(',') }.join('-')",
    "1,a-2,b"
);
ruby_test!(
    test_enumerator_lazy_chunk,
    "puts [1, 2, 2, 3].lazy.chunk { |x| x.even? }.force.map { |k, v| \"#{k}:#{v.join(',')}\" }.join('-')",
    "false:1-true:2,2-false:3"
);
