macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_date_civil_basic,
    "require 'date'; puts Date.civil(2001, 2, 3).to_s",
    "2001-02-03"
);
ruby_test!(
    test_date_civil_negative_year,
    "require 'date'; puts Date.civil(-2001, 2, 3).to_s",
    "-2001-02-03"
);
ruby_test!(
    test_date_civil_negative_month,
    "require 'date'; puts Date.civil(2001, -2, 3).to_s",
    "2001-11-03"
); // -2 is November
ruby_test!(
    test_date_civil_negative_day,
    "require 'date'; puts Date.civil(2001, 2, -3).to_s",
    "2001-02-26"
); // 3rd from end of Feb (28 days) -> 26
ruby_test!(
    test_date_civil_invalid_month_error,
    "require 'date'; begin; Date.civil(2001, 13, 3); rescue Date::Error; puts 'err'; end",
    "err"
);
ruby_test!(
    test_date_civil_invalid_day_error,
    "require 'date'; begin; Date.civil(2001, 2, 29); rescue Date::Error; puts 'err'; end",
    "err"
); // not leap
ruby_test!(
    test_date_civil_leap_day,
    "require 'date'; puts Date.civil(2004, 2, 29).to_s",
    "2004-02-29"
);
ruby_test!(
    test_date_new_alias,
    "require 'date'; puts Date.new(2001, 2, 3).to_s",
    "2001-02-03"
);
