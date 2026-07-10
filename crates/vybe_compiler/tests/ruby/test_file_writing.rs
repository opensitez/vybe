use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_file_writing_write, "File.write('/tmp/test_file_writing.txt', 'hello'); puts File.read('/tmp/test_file_writing.txt')", "hello");
ruby_test!(test_file_writing_write_length, "puts File.write('/tmp/test_file_writing.txt', 'hello')", "5");
ruby_test!(test_file_writing_write_offset, "File.write('/tmp/test_file_writing.txt', 'hello'); File.write('/tmp/test_file_writing.txt', 'a', 1); puts File.read('/tmp/test_file_writing.txt')", "hallo");
ruby_test!(test_file_writing_write_mode, "File.write('/tmp/test_file_writing.txt', 'a'); File.write('/tmp/test_file_writing.txt', 'b', mode: 'a'); puts File.read('/tmp/test_file_writing.txt')", "ab");
ruby_test!(test_file_writing_instance_write, "f = File.open('/tmp/test_file_writing.txt', 'w'); f.write('hello'); f.close; puts File.read('/tmp/test_file_writing.txt')", "hello");
ruby_test!(test_file_writing_instance_puts, "f = File.open('/tmp/test_file_writing.txt', 'w'); f.puts('hello'); f.close; puts File.read('/tmp/test_file_writing.txt')", "hello\\n");
ruby_test!(test_file_writing_instance_print, "f = File.open('/tmp/test_file_writing.txt', 'w'); f.print('hello'); f.close; puts File.read('/tmp/test_file_writing.txt')", "hello");
ruby_test!(test_file_writing_instance_flush, "f = File.open('/tmp/test_file_writing.txt', 'w'); f.write('hello'); puts f.flush.class.name; f.close", "File");
ruby_test!(test_file_writing_binwrite, "File.binwrite('/tmp/test_file_writing.txt', 'hello'); puts File.read('/tmp/test_file_writing.txt')", "hello");
