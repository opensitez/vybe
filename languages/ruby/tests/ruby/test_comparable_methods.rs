macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_comparable_comparison_lt,
    "class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(1) < C.new(2)",
    "true"
);
ruby_test!(
    test_comparable_comparison_lte,
    "class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(1) <= C.new(1)",
    "true"
);
ruby_test!(
    test_comparable_comparison_eq,
    "class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(1) == C.new(1)",
    "true"
);
ruby_test!(
    test_comparable_comparison_gt,
    "class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(2) > C.new(1)",
    "true"
);
ruby_test!(
    test_comparable_comparison_gte,
    "class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(2) >= C.new(2)",
    "true"
);
ruby_test!(
    test_comparable_between,
    "class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(2).between?(C.new(1), C.new(3))",
    "true"
);
ruby_test!(
    test_comparable_clamp,
    "class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(5).clamp(C.new(1), C.new(3)).val",
    "3"
);
ruby_test!(
    test_comparable_invalid_comparison,
    "class C; include Comparable; def <=>(o); nil; end; end; begin; C.new < C.new; rescue ArgumentError; puts 'err'; end",
    "err"
);
