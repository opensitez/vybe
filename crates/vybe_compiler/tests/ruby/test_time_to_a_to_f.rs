use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_time_to_f_zero, "puts Time.at(0).to_f", "0.0");
ruby_test!(test_time_to_f_fractional, "puts Time.at(1.5).to_f", "1.5");
ruby_test!(test_time_to_r_zero, "puts Time.at(0).to_r", "0/1");
ruby_test!(test_time_to_r_fractional, "puts Time.at(1.5).to_r", "3/2");
ruby_test!(test_time_to_i_basic, "puts Time.at(1.9).to_i", "1"); // truncates
ruby_test!(test_time_to_a_basic, "puts Time.utc(2000, 1, 1).to_a.length", "10"); // [sec, min, hour, day, month, year, wday, yday, isdst, zone]
ruby_test!(test_time_to_a_values, "a = Time.utc(2000, 1, 2, 3, 4, 5).to_a; puts \"#{a[0]}-#{a[1]}-#{a[2]}-#{a[3]}-#{a[4]}-#{a[5]}\"", "5-4-3-2-1-2000"); // sec-min-hour-day-mon-year
ruby_test!(test_time_to_a_dst_zone, "a = Time.utc(2000, 1, 1).to_a; puts \"#{a[8]}-#{a[9]}\"", "false-UTC");
