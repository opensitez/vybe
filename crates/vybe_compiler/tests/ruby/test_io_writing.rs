macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_io_writing_puts,
    "r, w = IO.pipe; w.puts('a', 'b'); w.close; puts r.read.chomp.split(\"\\n\").join('-')",
    "a-b"
);
ruby_test!(
    test_io_writing_print,
    "r, w = IO.pipe; w.print('a', 'b'); w.close; puts r.read",
    "ab"
);
ruby_test!(
    test_io_writing_printf,
    "r, w = IO.pipe; w.printf('%d %s', 1, 'a'); w.close; puts r.read",
    "1 a"
);
ruby_test!(
    test_io_writing_putc,
    "r, w = IO.pipe; w.putc(65); w.close; puts r.read",
    "A"
);
ruby_test!(
    test_io_writing_write,
    "r, w = IO.pipe; puts w.write('abc')",
    "3"
);
ruby_test!(
    test_io_writing_flush,
    "r, w = IO.pipe; puts w.flush == w",
    "true"
);
ruby_test!(
    test_io_writing_sync,
    "r, w = IO.pipe; w.sync = true; puts w.sync",
    "true"
);
ruby_test!(
    test_io_writing_fsync,
    "r, w = IO.pipe; begin; w.fsync; rescue IOError, NotImplementedError, Errno::EINVAL; puts 'err'; end",
    "err"
); // fsync on pipe usually raises EINVAL or IOError
