use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dir_home_basic, "puts Dir.home.start_with?('/')", "true");
ruby_test!(test_dir_home_current_user, "puts Dir.home == ENV['HOME']", "true");
ruby_test!(test_dir_home_other_user, "begin; Dir.home('non_existent_user'); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_dir_home_root_user, "puts Dir.home('root').start_with?('/')", "true");
