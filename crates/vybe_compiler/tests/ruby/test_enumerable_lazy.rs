macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_enumerable_lazy_basic,
    "puts [1, 2, 3].lazy.class.name",
    "Enumerator::Lazy"
);
ruby_test!(
    test_enumerable_lazy_map,
    "puts [1, 2, 3].lazy.map { |x| x * 10 }.first(2).join('-')",
    "10-20"
);
ruby_test!(
    test_enumerable_lazy_select,
    "puts [1, 2, 3, 4].lazy.select { |x| x.even? }.first(1).join('-')",
    "2"
);
ruby_test!(
    test_enumerable_lazy_infinite,
    "puts (1..Float::INFINITY).lazy.map { |x| x * 2 }.first(3).join('-')",
    "2-4-6"
);
ruby_test!(
    test_enumerable_lazy_force,
    "puts [1, 2, 3].lazy.map { |x| x * 10 }.force.join('-')",
    "10-20-30"
);
ruby_test!(
    test_enumerable_lazy_drop,
    "puts (1..10).lazy.drop(2).first(2).join('-')",
    "3-4"
);
ruby_test!(
    test_enumerable_lazy_take,
    "puts (1..10).lazy.take(3).force.join('-')",
    "1-2-3"
);
ruby_test!(
    test_enumerable_lazy_grep,
    "puts (1..10).lazy.grep(3..6).first(2).join('-')",
    "3-4"
);
ruby_test!(
    test_enumerable_lazy_reject,
    "puts (1..10).lazy.reject { |x| x.even? }.first(2).join('-')",
    "1-3"
);
ruby_test!(
    test_enumerable_lazy_zip,
    "puts (1..3).lazy.zip(['a', 'b', 'c']).first(2).map{|a| a.join}.join('-')",
    "1a-2b"
);
