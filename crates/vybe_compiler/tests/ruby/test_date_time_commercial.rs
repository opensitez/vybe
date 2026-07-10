use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_date_commercial_basic, "require 'date'; puts Date.commercial(2001, 5, 6).to_s", "2001-02-03"); // 2001-W05-6
ruby_test!(test_date_commercial_default_day, "require 'date'; puts Date.commercial(2001, 5).to_s", "2001-01-29"); // defaults to Monday (1)
ruby_test!(test_date_commercial_default_week, "require 'date'; puts Date.commercial(2001).to_s", "2001-01-01"); // defaults to week 1, day 1
ruby_test!(test_date_commercial_negative_week, "require 'date'; puts Date.commercial(2001, -1).to_s", "2001-12-31"); // last week
ruby_test!(test_date_commercial_negative_day, "require 'date'; puts Date.commercial(2001, 5, -1).to_s", "2001-02-04"); // last day of week (Sunday)
ruby_test!(test_date_commercial_invalid_week_error, "require 'date'; begin; Date.commercial(2001, 54); rescue Date::Error; puts 'err'; end", "err");
ruby_test!(test_date_commercial_invalid_day_error, "require 'date'; begin; Date.commercial(2001, 5, 8); rescue Date::Error; puts 'err'; end", "err");
ruby_test!(test_date_cwyear, "require 'date'; puts Date.new(2001, 2, 3).cwyear", "2001");
ruby_test!(test_date_cweek, "require 'date'; puts Date.new(2001, 2, 3).cweek", "5");
ruby_test!(test_date_cwday, "require 'date'; puts Date.new(2001, 2, 3).cwday", "6"); // Saturday
