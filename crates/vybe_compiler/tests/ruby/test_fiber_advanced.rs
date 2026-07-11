
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_fiber_advanced_transfer, "require 'fiber'; f1 = Fiber.new { |x| x * 2 }; puts f1.transfer(21)", "42");
ruby_test!(test_fiber_advanced_yield_args, "f = Fiber.new { |x| Fiber.yield(x + 1) }; puts f.resume(10)", "11");
ruby_test!(test_fiber_advanced_resume_args, "f = Fiber.new { |x| x + 1 }; puts f.resume(10)", "11");
ruby_test!(test_fiber_advanced_current, "puts Fiber.current.class.name", "Fiber");
ruby_test!(test_fiber_advanced_raise, "f = Fiber.new { Fiber.yield }; begin; f.raise 'err'; rescue => e; puts e.message; end", "err");
