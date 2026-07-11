
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_fiber_scheduler, "puts Fiber.scheduler.nil?", "true");
ruby_test!(test_fiber_set_scheduler_invalid, "begin; Fiber.set_scheduler(Object.new); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_fiber_current, "puts Fiber.current.class.name", "Fiber");
ruby_test!(test_fiber_yield_args, "f = Fiber.new { |a| Fiber.yield(a * 2) }; puts f.resume(21)", "42");
ruby_test!(test_fiber_yield_multiple, "f = Fiber.new { Fiber.yield(1, 2) }; puts f.resume.join('-')", "1-2");
ruby_test!(test_fiber_resume_args, "f = Fiber.new { |a, b| a + b }; puts f.resume(1, 2)", "3");
