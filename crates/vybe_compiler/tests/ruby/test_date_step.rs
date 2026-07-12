macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_date_step_basic,
    "require 'date'; acc = []; Date.new(2000,1,1).step(Date.new(2000,1,3)) {|d| acc << d.day}; puts acc.join('-')",
    "1-2-3"
);
ruby_test!(
    test_date_step_negative,
    "require 'date'; acc = []; Date.new(2000,1,3).step(Date.new(2000,1,1), -1) {|d| acc << d.day}; puts acc.join('-')",
    "3-2-1"
);
ruby_test!(
    test_date_step_no_block,
    "require 'date'; puts Date.new(2000,1,1).step(Date.new(2000,1,3)).is_a?(Enumerator)",
    "true"
);
ruby_test!(
    test_date_upto,
    "require 'date'; acc = []; Date.new(2000,1,1).upto(Date.new(2000,1,3)) {|d| acc << d.day}; puts acc.join('-')",
    "1-2-3"
);
ruby_test!(
    test_date_downto,
    "require 'date'; acc = []; Date.new(2000,1,3).downto(Date.new(2000,1,1)) {|d| acc << d.day}; puts acc.join('-')",
    "3-2-1"
);
