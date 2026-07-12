macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_dir_methods_pwd, "puts Dir.pwd.class.name", "String");
ruby_test!(
    test_dir_methods_getwd,
    "puts Dir.getwd.class.name",
    "String"
);
ruby_test!(
    test_dir_methods_mkdir_rmdir,
    "Dir.mkdir('/tmp/test_dir_methods'); puts Dir.exist?('/tmp/test_dir_methods'); Dir.rmdir('/tmp/test_dir_methods'); puts Dir.exist?('/tmp/test_dir_methods')",
    "true\\nfalse"
);
ruby_test!(
    test_dir_methods_entries,
    "Dir.mkdir('/tmp/test_dir_methods_entries'); File.write('/tmp/test_dir_methods_entries/a', 'a'); puts Dir.entries('/tmp/test_dir_methods_entries').sort.join('-'); File.delete('/tmp/test_dir_methods_entries/a'); Dir.rmdir('/tmp/test_dir_methods_entries')",
    ".-..-a"
);
ruby_test!(
    test_dir_methods_foreach,
    "Dir.mkdir('/tmp/test_dir_methods_foreach'); File.write('/tmp/test_dir_methods_foreach/a', 'a'); acc = []; Dir.foreach('/tmp/test_dir_methods_foreach') { |e| acc << e }; puts acc.sort.join('-'); File.delete('/tmp/test_dir_methods_foreach/a'); Dir.rmdir('/tmp/test_dir_methods_foreach')",
    ".-..-a"
);
ruby_test!(
    test_dir_methods_glob,
    "Dir.mkdir('/tmp/test_dir_methods_glob'); File.write('/tmp/test_dir_methods_glob/a.rb', 'a'); puts Dir.glob('/tmp/test_dir_methods_glob/*.rb').length; File.delete('/tmp/test_dir_methods_glob/a.rb'); Dir.rmdir('/tmp/test_dir_methods_glob')",
    "1"
);
ruby_test!(test_dir_methods_home, "puts Dir.home.class.name", "String");
ruby_test!(
    test_dir_methods_empty,
    "Dir.mkdir('/tmp/test_dir_methods_empty'); puts Dir.empty?('/tmp/test_dir_methods_empty'); Dir.rmdir('/tmp/test_dir_methods_empty')",
    "true"
);
ruby_test!(test_dir_methods_exist, "puts Dir.exist?('/dev')", "true");
ruby_test!(test_dir_methods_exists, "puts Dir.exists?('/dev')", "true"); // deprecated but works
