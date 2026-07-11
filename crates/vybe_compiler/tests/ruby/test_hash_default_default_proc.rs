
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_default_get, "h = Hash.new(5); puts h.default", "5");
ruby_test!(test_default_get_key, "h = Hash.new(5); puts h.default(:a)", "5");
ruby_test!(test_default_set, "h = Hash.new; h.default = 5; puts h[:a]", "5");
ruby_test!(test_default_set_overwrites_block, "h = Hash.new {|hash, key| 1}; h.default = 5; puts h[:a]", "5"); // removes block
ruby_test!(test_default_proc_get, "h = Hash.new {|hash, key| 1}; puts h.default_proc.is_a?(Proc)", "true");
ruby_test!(test_default_proc_get_none, "h = Hash.new(5); puts h.default_proc.nil?", "true");
ruby_test!(test_default_proc_set, "h = Hash.new(5); h.default_proc = proc {|hash, key| 1}; puts h[:a]", "1");
ruby_test!(test_default_proc_set_nil, "h = Hash.new {|hash, key| 1}; h.default_proc = nil; puts h[:a].nil?", "true"); // removes block, default becomes nil
ruby_test!(test_default_proc_set_not_proc, "h = Hash.new; begin; h.default_proc = 5; rescue TypeError; puts 'err'; end", "err");
ruby_test!(test_default_proc_set_lambda, "h = Hash.new; h.default_proc = ->(hash, key) { 1 }; puts h[:a]", "1");
