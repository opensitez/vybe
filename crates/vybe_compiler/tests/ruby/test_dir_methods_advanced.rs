
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dir_children, "puts Dir.children('.').class.name", "Array");
ruby_test!(test_dir_each_child, "acc = []; Dir.each_child('.') { |f| acc << f if f == 'Cargo.toml' }; puts acc.join", "Cargo.toml");
ruby_test!(test_dir_empty, "Dir.mkdir('test_empty_dir'); puts Dir.empty?('test_empty_dir'); Dir.rmdir('test_empty_dir')", "true");
ruby_test!(test_dir_exist, "puts Dir.exist?('.')", "true");
ruby_test!(test_dir_home, "puts Dir.home.class.name", "String");
ruby_test!(test_dir_pwd, "puts Dir.pwd.class.name", "String");
ruby_test!(test_dir_getwd, "puts Dir.getwd == Dir.pwd", "true");
ruby_test!(test_dir_glob, "puts Dir.glob('*.toml').include?('Cargo.toml').to_s", "true");
ruby_test!(test_dir_open, "d = Dir.open('.'); puts d.class.name; d.close", "Dir");
