macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_file_path_basename,
    "puts File.basename('/foo/bar.txt')",
    "bar.txt"
);
ruby_test!(
    test_file_path_basename_suffix,
    "puts File.basename('/foo/bar.txt', '.txt')",
    "bar"
);
ruby_test!(
    test_file_path_dirname,
    "puts File.dirname('/foo/bar.txt')",
    "/foo"
);
ruby_test!(
    test_file_path_extname,
    "puts File.extname('/foo/bar.txt')",
    ".txt"
);
ruby_test!(
    test_file_path_extname_empty,
    "puts File.extname('/foo/bar')",
    ""
);
ruby_test!(
    test_file_path_join,
    "puts File.join('foo', 'bar', 'baz')",
    "foo/bar/baz"
);
ruby_test!(
    test_file_path_split,
    "puts File.split('/foo/bar.txt').join('-')",
    "/foo-bar.txt"
);
ruby_test!(
    test_file_path_expand_path,
    "puts File.expand_path('bar', '/foo')",
    "/foo/bar"
);
ruby_test!(
    test_file_path_absolute,
    "puts File.absolute_path?('/foo')",
    "true"
);
ruby_test!(
    test_file_path_absolute_false,
    "puts File.absolute_path?('foo')",
    "false"
);
ruby_test!(
    test_file_path_fnmatch,
    "puts File.fnmatch('*.txt', 'foo.txt')",
    "true"
);
ruby_test!(
    test_file_path_fnmatch_false,
    "puts File.fnmatch('*.txt', 'foo.rb')",
    "false"
);
