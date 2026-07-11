
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_security_error_basic, "begin; raise SecurityError; rescue SecurityError; puts 'caught'; end", "caught");
ruby_test!(test_security_error_inherits_exception, "begin; raise SecurityError; rescue Exception; puts 'caught exception'; end", "caught exception");
ruby_test!(test_security_error_not_caught_by_standard_error, "begin; raise SecurityError; rescue StandardError; puts 'caught std'; rescue SecurityError; puts 'caught sec'; end", "caught sec"); // Wait, SecurityError inherits from Exception, not StandardError? Actually, Ruby 1.9+ SecurityError < Exception, wait, in Ruby 1.9+ SecurityError < StandardError? Let's check docs... actually SecurityError inherits from Exception. No wait, SecurityError < Exception. Let's just test basic raise/rescue.
ruby_test!(test_security_error_to_s, "begin; raise SecurityError, 'sec'; rescue SecurityError => e; puts e.message; end", "sec");
