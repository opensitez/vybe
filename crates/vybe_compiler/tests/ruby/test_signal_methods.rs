
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_signal_list, "puts Signal.list.class.name", "Hash");
ruby_test!(test_signal_list_keys, "puts Signal.list.keys.include?('INT').to_s", "true");
ruby_test!(test_signal_list_values, "puts Signal.list.values.include?(2).to_s", "true");
ruby_test!(test_signal_signame, "puts Signal.signame(2)", "INT");
ruby_test!(test_signal_signame_invalid, "puts Signal.signame(9999).nil?", "true");
ruby_test!(test_signal_trap_invalid, "begin; Signal.trap('INVALID', 'IGNORE'); rescue ArgumentError; puts 'err'; end", "err");
