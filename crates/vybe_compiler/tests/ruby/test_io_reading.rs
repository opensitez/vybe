
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_io_reading_gets, "r, w = IO.pipe; w.puts('hello'); w.close; puts r.gets", "hello");
ruby_test!(test_io_reading_read, "r, w = IO.pipe; w.write('hello'); w.close; puts r.read", "hello");
ruby_test!(test_io_reading_read_length, "r, w = IO.pipe; w.write('hello'); w.close; puts r.read(3)", "hel");
ruby_test!(test_io_reading_readpartial, "r, w = IO.pipe; w.write('hello'); w.close; puts r.readpartial(10)", "hello");
ruby_test!(test_io_reading_getc, "r, w = IO.pipe; w.write('hello'); w.close; puts r.getc", "h");
ruby_test!(test_io_reading_getbyte, "r, w = IO.pipe; w.write('A'); w.close; puts r.getbyte", "65");
ruby_test!(test_io_reading_readlines, "r, w = IO.pipe; w.puts('a'); w.puts('b'); w.close; puts r.readlines.map(&:chomp).join('-')", "a-b");
ruby_test!(test_io_reading_each_line, "r, w = IO.pipe; w.puts('a'); w.puts('b'); w.close; acc = []; r.each_line { |l| acc << l.chomp }; puts acc.join('-')", "a-b");
ruby_test!(test_io_reading_eof, "r, w = IO.pipe; w.close; puts r.eof?", "true");
ruby_test!(test_io_reading_ungetc, "r, w = IO.pipe; w.write('ello'); w.close; r.ungetc('h'); puts r.read", "hello");
ruby_test!(test_io_reading_ungetbyte, "r, w = IO.pipe; w.write('B'); w.close; r.ungetbyte(65); puts r.read", "AB");
