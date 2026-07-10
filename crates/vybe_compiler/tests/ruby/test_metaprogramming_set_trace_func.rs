use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_set_trace_func_basic, "acc = []; set_trace_func proc { |event, file, line, id, binding, classname| acc << event }; def foo; end; foo; set_trace_func nil; puts acc.include?('call')", "true");
ruby_test!(test_set_trace_func_return, "acc = []; set_trace_func proc { |event, file, line, id, binding, classname| acc << event }; def foo; end; foo; set_trace_func nil; puts acc.include?('return')", "true");
ruby_test!(test_set_trace_func_line, "acc = []; set_trace_func proc { |event, file, line, id, binding, classname| acc << event }; x = 1; set_trace_func nil; puts acc.include?('line')", "true");
