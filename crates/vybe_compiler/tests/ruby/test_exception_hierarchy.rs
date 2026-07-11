
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_exception_hierarchy_standard, "puts StandardError.ancestors.include?(Exception).to_s", "true");
ruby_test!(test_exception_hierarchy_argument, "puts ArgumentError.ancestors.include?(StandardError).to_s", "true");
ruby_test!(test_exception_hierarchy_name, "puts NameError.ancestors.include?(StandardError).to_s", "true");
ruby_test!(test_exception_hierarchy_no_method, "puts NoMethodError.ancestors.include?(NameError).to_s", "true");
ruby_test!(test_exception_hierarchy_runtime, "puts RuntimeError.ancestors.include?(StandardError).to_s", "true");
ruby_test!(test_exception_hierarchy_type, "puts TypeError.ancestors.include?(StandardError).to_s", "true");
ruby_test!(test_exception_hierarchy_index, "puts IndexError.ancestors.include?(StandardError).to_s", "true");
ruby_test!(test_exception_hierarchy_key, "puts KeyError.ancestors.include?(IndexError).to_s", "true");
ruby_test!(test_exception_hierarchy_range, "puts RangeError.ancestors.include?(StandardError).to_s", "true");
ruby_test!(test_exception_hierarchy_float_domain, "puts FloatDomainError.ancestors.include?(RangeError).to_s", "true");
ruby_test!(test_exception_hierarchy_system_exit, "puts SystemExit.ancestors.include?(Exception).to_s", "true");
ruby_test!(test_exception_hierarchy_interrupt, "puts Interrupt.ancestors.include?(SignalException).to_s", "true");
ruby_test!(test_exception_message, "puts Exception.new('foo').message", "foo");
ruby_test!(test_exception_to_s, "puts Exception.new('foo').to_s", "foo");
ruby_test!(test_exception_inspect, "puts Exception.new('foo').inspect", "#<Exception: foo>");
ruby_test!(test_exception_backtrace, "begin; raise 'foo'; rescue => e; puts e.backtrace.class.name; end", "Array");
ruby_test!(test_exception_set_backtrace, "e = Exception.new; e.set_backtrace(['a', 'b']); puts e.backtrace.join('-')", "a-b");
