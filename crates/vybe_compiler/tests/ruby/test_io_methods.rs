
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_io_methods_pipe, "r, w = IO.pipe; w.write('hello'); w.close; puts r.read; r.close", "hello");
ruby_test!(test_io_methods_select, "r, w = IO.pipe; w.write('a'); puts IO.select([r], nil, nil, 0).length; w.close; r.close", "1");
ruby_test!(test_io_methods_popen, "f = IO.popen('echo hello'); puts f.read; f.close", "hello\\n");
ruby_test!(test_io_methods_readlines, "File.write('/tmp/test_io_methods.txt', \"a\\nb\\nc\"); puts IO.readlines('/tmp/test_io_methods.txt').join('-')", "a\\n-b\\n-c");
ruby_test!(test_io_methods_read, "File.write('/tmp/test_io_methods.txt', 'hello'); puts IO.read('/tmp/test_io_methods.txt')", "hello");
ruby_test!(test_io_methods_write, "IO.write('/tmp/test_io_methods.txt', 'hello'); puts IO.read('/tmp/test_io_methods.txt')", "hello");
ruby_test!(test_io_methods_binread, "File.write('/tmp/test_io_methods.txt', 'hello'); puts IO.binread('/tmp/test_io_methods.txt')", "hello");
ruby_test!(test_io_methods_binwrite, "IO.binwrite('/tmp/test_io_methods.txt', 'hello'); puts IO.read('/tmp/test_io_methods.txt')", "hello");
ruby_test!(test_io_methods_copy_stream, "File.write('/tmp/test_io_methods.txt', 'hello'); IO.copy_stream('/tmp/test_io_methods.txt', '/tmp/test_io_methods_out.txt'); puts IO.read('/tmp/test_io_methods_out.txt')", "hello");
