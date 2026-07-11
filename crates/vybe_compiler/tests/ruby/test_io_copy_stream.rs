
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_copy_stream_basic, "require 'tempfile'; t1 = Tempfile.new('src'); t2 = Tempfile.new('dst'); t1.write('hello'); t1.rewind; IO.copy_stream(t1, t2); t2.rewind; puts t2.read", "hello");
ruby_test!(test_copy_stream_length, "require 'tempfile'; t1 = Tempfile.new('src'); t2 = Tempfile.new('dst'); t1.write('hello'); t1.rewind; IO.copy_stream(t1, t2, 3); t2.rewind; puts t2.read", "hel");
ruby_test!(test_copy_stream_offset, "require 'tempfile'; t1 = Tempfile.new('src'); t2 = Tempfile.new('dst'); t1.write('hello'); IO.copy_stream(t1.path, t2, nil, 2); t2.rewind; puts t2.read", "llo"); // reads from path with offset
ruby_test!(test_copy_stream_return_value, "require 'tempfile'; t1 = Tempfile.new('src'); t2 = Tempfile.new('dst'); t1.write('hello'); t1.rewind; puts IO.copy_stream(t1, t2)", "5"); // returns bytes copied
ruby_test!(test_copy_stream_error_not_exist, "require 'tempfile'; t2 = Tempfile.new('dst'); begin; IO.copy_stream('non_existent_file.txt', t2); rescue Errno::ENOENT; puts 'err'; end", "err");
