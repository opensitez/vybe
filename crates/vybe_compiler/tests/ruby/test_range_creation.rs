use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_range_creation_inclusive, "r = 1..5; puts r.class.name", "Range");
ruby_test!(test_range_creation_exclusive, "r = 1...5; puts r.class.name", "Range");
ruby_test!(test_range_creation_endless, "r = (1..); puts r.class.name", "Range");
ruby_test!(test_range_creation_beginless, "r = (..5); puts r.class.name", "Range");
ruby_test!(test_range_begin, "puts (1..5).begin", "1");
ruby_test!(test_range_end, "puts (1..5).end", "5");
ruby_test!(test_range_exclude_end_inclusive, "puts (1..5).exclude_end?", "false");
ruby_test!(test_range_exclude_end_exclusive, "puts (1...5).exclude_end?", "true");
ruby_test!(test_range_eq, "puts (1..5) == (1..5)", "true");
ruby_test!(test_range_not_eq, "puts (1..5) == (1...5)", "false");
ruby_test!(test_range_eql, "puts (1..5).eql?(1..5)", "true");
ruby_test!(test_range_hash, "puts (1..5).hash == (1..5).hash", "true");
