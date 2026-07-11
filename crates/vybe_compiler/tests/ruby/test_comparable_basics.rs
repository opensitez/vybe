
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_comparable_less_than, "class A; include Comparable; attr_reader :x; def initialize(x); @x = x; end; def <=>(other); @x <=> other.x; end; end; puts A.new(1) < A.new(2)", "true");
ruby_test!(test_comparable_greater_than, "class A; include Comparable; attr_reader :x; def initialize(x); @x = x; end; def <=>(other); @x <=> other.x; end; end; puts A.new(1) > A.new(2)", "false");
ruby_test!(test_comparable_less_eq, "class A; include Comparable; attr_reader :x; def initialize(x); @x = x; end; def <=>(other); @x <=> other.x; end; end; puts A.new(1) <= A.new(1)", "true");
ruby_test!(test_comparable_greater_eq, "class A; include Comparable; attr_reader :x; def initialize(x); @x = x; end; def <=>(other); @x <=> other.x; end; end; puts A.new(2) >= A.new(1)", "true");
ruby_test!(test_comparable_equal, "class A; include Comparable; attr_reader :x; def initialize(x); @x = x; end; def <=>(other); @x <=> other.x; end; end; puts A.new(1) == A.new(1)", "true");
