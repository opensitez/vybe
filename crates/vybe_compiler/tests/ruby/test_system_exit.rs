
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_system_exit_status, "begin; exit(42); rescue SystemExit => e; puts e.status; end", "42");
ruby_test!(test_system_exit_success_true, "begin; exit(0); rescue SystemExit => e; puts e.success?; end", "true");
ruby_test!(test_system_exit_success_false, "begin; exit(1); rescue SystemExit => e; puts e.success?; end", "false");
ruby_test!(test_system_exit_default_status, "begin; exit; rescue SystemExit => e; puts e.status; end", "0");
ruby_test!(test_system_exit_abort, "begin; abort('err'); rescue SystemExit => e; puts \"#{e.status}-#{e.message}\"; end", "1-err");
