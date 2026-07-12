macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_proc_source_location_basic,
    "p = Proc.new { }; puts p.source_location[0].end_with?('.rb') || p.source_location[0] == '-e'",
    "true"
);
ruby_test!(
    test_proc_source_location_line,
    "\np = Proc.new { }; puts p.source_location[1]",
    "2"
);
ruby_test!(
    test_lambda_source_location_line,
    "\n\nl = lambda { }; puts l.source_location[1]",
    "3"
);
ruby_test!(
    test_proc_source_location_native,
    "p = :to_s.to_proc; puts p.source_location.nil?",
    "true"
); // Native procs like Symbol#to_proc return nil wait no, in some Rubies it might return something else? Actually Symbol#to_proc returns nil for source_location in MRI.
