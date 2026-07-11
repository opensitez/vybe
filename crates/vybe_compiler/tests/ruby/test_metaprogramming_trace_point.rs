
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_trace_point_basic, "acc = []; tp = TracePoint.new(:call) do |t| acc << t.method_id if t.method_id == :foo end; def foo; end; tp.enable; foo; tp.disable; puts acc.include?(:foo)", "true");
ruby_test!(test_trace_point_line, "acc = []; tp = TracePoint.new(:line) do |t| acc << t.line end; tp.enable; x = 1; tp.disable; puts acc.size > 0", "true");
ruby_test!(test_trace_point_return, "acc = []; tp = TracePoint.new(:return) do |t| acc << t.return_value if t.method_id == :foo end; def foo; 'ret'; end; tp.enable; foo; tp.disable; puts acc.include?('ret')", "true");
ruby_test!(test_trace_point_raise, "acc = []; tp = TracePoint.new(:raise) do |t| acc << t.raised_exception.class end; tp.enable; begin; raise 'err'; rescue; end; tp.disable; puts acc.include?(RuntimeError)", "true");
