
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_to_proc_basic, "h = {a: 1, b: 2}; p = h.to_proc; puts p.call(:a)", "1");
ruby_test!(test_to_proc_missing, "h = {a: 1}; p = h.to_proc; puts p.call(:b).nil?", "true");
ruby_test!(test_to_proc_map, "h = {a: 1, b: 2, c: 3}; puts [:a, :b, :c].map(&h).join('-')", "1-2-3");
ruby_test!(test_to_proc_returns_proc, "puts {a: 1}.to_proc.is_a?(Proc)", "true");
ruby_test!(test_to_proc_arity, "puts {a: 1}.to_proc.arity", "1"); // Hash to_proc takes 1 argument
ruby_test!(test_to_proc_default_value, "h = Hash.new('def'); p = h.to_proc; puts p.call(:a)", "def");
ruby_test!(test_to_proc_default_proc, "h = Hash.new {|hash, key| 'def'}; p = h.to_proc; puts p.call(:a)", "def");
ruby_test!(test_to_proc_nil_key, "h = {nil => 1}; puts h.to_proc.call(nil)", "1");
ruby_test!(test_to_proc_array_key, "h = {[1, 2] => 3}; puts h.to_proc.call([1, 2])", "3");
