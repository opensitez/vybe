use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_method_added, "class A; @acc = []; def self.method_added(m); @acc << m unless m == :method_added || m == :acc; end; def foo; end; def self.acc; @acc; end; end; puts A.acc.include?(:foo)", "true");
ruby_test!(test_method_removed, "class A; @acc = []; def self.method_removed(m); @acc << m; end; def foo; end; remove_method :foo; def self.acc; @acc; end; end; puts A.acc.include?(:foo)", "true");
ruby_test!(test_method_undefined, "class A; @acc = []; def self.method_undefined(m); @acc << m; end; def foo; end; undef_method :foo; def self.acc; @acc; end; end; puts A.acc.include?(:foo)", "true");
