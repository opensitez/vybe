use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_file_reading_read, "File.write('/tmp/test_file_reading.txt', 'hello'); puts File.read('/tmp/test_file_reading.txt')", "hello");
ruby_test!(test_file_reading_read_length, "File.write('/tmp/test_file_reading.txt', 'hello'); puts File.read('/tmp/test_file_reading.txt', 2)", "he");
ruby_test!(test_file_reading_read_offset, "File.write('/tmp/test_file_reading.txt', 'hello'); puts File.read('/tmp/test_file_reading.txt', 2, 1)", "el");
ruby_test!(test_file_reading_readlines, "File.write('/tmp/test_file_reading.txt', \"a\\nb\\nc\"); puts File.readlines('/tmp/test_file_reading.txt').join('-')", "a\\n-b\\n-c");
ruby_test!(test_file_reading_readlines_chomp, "File.write('/tmp/test_file_reading.txt', \"a\\nb\\nc\"); puts File.readlines('/tmp/test_file_reading.txt', chomp: true).join('-')", "a-b-c");
ruby_test!(test_file_reading_instance_read, "File.write('/tmp/test_file_reading.txt', 'hello'); f = File.open('/tmp/test_file_reading.txt'); puts f.read; f.close", "hello");
ruby_test!(test_file_reading_instance_gets, "File.write('/tmp/test_file_reading.txt', \"a\\nb\"); f = File.open('/tmp/test_file_reading.txt'); puts f.gets; f.close", "a\\n");
ruby_test!(test_file_reading_instance_each_line, "File.write('/tmp/test_file_reading.txt', \"a\\nb\"); acc = []; File.open('/tmp/test_file_reading.txt') { |f| f.each_line { |l| acc << l.chomp } }; puts acc.join('-')", "a-b");
ruby_test!(test_file_reading_instance_eof, "File.write('/tmp/test_file_reading.txt', ''); f = File.open('/tmp/test_file_reading.txt'); puts f.eof?; f.close", "true");
ruby_test!(test_file_reading_instance_pos, "File.write('/tmp/test_file_reading.txt', 'hello'); f = File.open('/tmp/test_file_reading.txt'); f.read(2); puts f.pos; f.close", "2");
ruby_test!(test_file_reading_instance_rewind, "File.write('/tmp/test_file_reading.txt', 'hello'); f = File.open('/tmp/test_file_reading.txt'); f.read(2); f.rewind; puts f.pos; f.close", "0");
ruby_test!(test_file_reading_binread, "File.write('/tmp/test_file_reading.txt', 'hello'); puts File.binread('/tmp/test_file_reading.txt')", "hello");
