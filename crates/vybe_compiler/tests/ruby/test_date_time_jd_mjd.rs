
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_date_jd_basic, "require 'date'; puts Date.jd(2451944).to_s", "2001-02-03");
ruby_test!(test_date_jd_negative, "require 'date'; puts Date.jd(-1).to_s", "-4713-12-31");
ruby_test!(test_date_jd_zero, "require 'date'; puts Date.jd(0).to_s", "-4712-01-01"); // Julian 0
ruby_test!(test_date_mjd_basic, "require 'date'; puts Date.mjd(51943).to_s", "2001-02-03"); // mjd = jd - 2400001
ruby_test!(test_date_mjd_zero, "require 'date'; puts Date.mjd(0).to_s", "1858-11-17");
ruby_test!(test_date_jd_method, "require 'date'; puts Date.new(2001, 2, 3).jd", "2451944");
ruby_test!(test_date_mjd_method, "require 'date'; puts Date.new(2001, 2, 3).mjd", "51943");
ruby_test!(test_date_amjd_method, "require 'date'; puts Date.new(2001, 2, 3).amjd", "51943");
