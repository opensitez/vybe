macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_file_utime_basic,
    "require 'tempfile'; t = Tempfile.new('utime'); time = Time.now - 1000; puts File.utime(time, time, t.path)",
    "1"
);
ruby_test!(
    test_file_utime_multiple,
    "require 'tempfile'; t1 = Tempfile.new('utime1'); t2 = Tempfile.new('utime2'); time = Time.now; puts File.utime(time, time, t1.path, t2.path)",
    "2"
);
ruby_test!(
    test_file_utime_integer,
    "require 'tempfile'; t = Tempfile.new('utime'); time = Time.now.to_i; puts File.utime(time, time, t.path)",
    "1"
);
ruby_test!(
    test_file_utime_float,
    "require 'tempfile'; t = Tempfile.new('utime'); time = Time.now.to_f; puts File.utime(time, time, t.path)",
    "1"
);
ruby_test!(
    test_file_utime_error,
    "begin; File.utime(Time.now, Time.now, 'non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end",
    "err"
);
ruby_test!(
    test_file_utime_verify,
    "require 'tempfile'; t = Tempfile.new('utime'); time = Time.at(1000); File.utime(time, time, t.path); puts File.stat(t.path).mtime.to_i",
    "1000"
);
