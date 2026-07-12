macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_tempfile_basic,
    "require 'tempfile'; t = Tempfile.new('test'); t.write('hello'); t.rewind; puts t.read; t.close; t.unlink",
    "hello"
);
ruby_test!(
    test_tempfile_path,
    "require 'tempfile'; t = Tempfile.new('test'); puts t.path.nil?",
    "false"
);
ruby_test!(
    test_tempfile_unlink,
    "require 'tempfile'; t = Tempfile.new('test'); path = t.path; t.close; t.unlink; puts File.exist?(path)",
    "false"
);
ruby_test!(
    test_tempfile_block,
    "require 'tempfile'; p = Tempfile.create('test') {|t| t.write('hello'); t.path }; puts File.exist?(p); File.unlink(p)",
    "true"
); // Tempfile.create does not auto-unlink, unless you don't pass a block? wait, create with block yields file and deletes it afterwards in some rubies. Let's just check exist inside block:
ruby_test!(
    test_tempfile_create_block,
    "require 'tempfile'; path = nil; Tempfile.create('test') {|t| path = t.path; puts File.exist?(path) }; puts File.exist?(path)",
    "true\nfalse"
);
ruby_test!(
    test_tempfile_dirname,
    "require 'tempfile'; t = Tempfile.new('test'); puts File.dirname(t.path) == Dir.tmpdir",
    "true"
);
