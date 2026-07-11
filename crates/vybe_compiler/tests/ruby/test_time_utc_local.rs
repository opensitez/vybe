
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_time_utc_conversion, "t = Time.local(2000, 1, 1); t.utc; puts t.utc?", "true");
ruby_test!(test_time_getutc, "t = Time.local(2000, 1, 1); u = t.getutc; puts \"#{t.utc?}-#{u.utc?}\"", "false-true");
ruby_test!(test_time_localtime_conversion, "t = Time.utc(2000, 1, 1); t.localtime; puts t.utc?", "false");
ruby_test!(test_time_getlocal, "t = Time.utc(2000, 1, 1); l = t.getlocal; puts \"#{t.utc?}-#{l.utc?}\"", "true-false");
ruby_test!(test_time_getlocal_offset, "t = Time.utc(2000, 1, 1); l = t.getlocal('+09:00'); puts l.gmtoff", "32400"); // 9 hours in seconds
ruby_test!(test_time_localtime_offset, "t = Time.utc(2000, 1, 1); t.localtime('+09:00'); puts t.gmtoff", "32400");
