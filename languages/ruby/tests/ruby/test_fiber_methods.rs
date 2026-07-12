macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_fiber_creation,
    "f = Fiber.new { 42 }; puts f.resume",
    "42"
);
ruby_test!(
    test_fiber_yield,
    "f = Fiber.new { Fiber.yield 1; 2 }; puts \"#{f.resume}-#{f.resume}\"",
    "1-2"
);
ruby_test!(
    test_fiber_pass_args,
    "f = Fiber.new { |x| Fiber.yield x * 2; 3 }; puts \"#{f.resume(10)}-#{f.resume}\"",
    "20-3"
);
ruby_test!(
    test_fiber_alive,
    "f = Fiber.new { Fiber.yield 1 }; puts f.alive?; f.resume; puts f.alive?",
    "true\\ntrue"
); // Wait, it's alive until it finishes execution.
ruby_test!(
    test_fiber_alive_false,
    "f = Fiber.new { 1 }; f.resume; puts f.alive?",
    "false"
);
ruby_test!(test_fiber_current, "puts Fiber.current.class.name", "Fiber");
ruby_test!(
    test_fiber_dead_resume,
    "begin; f = Fiber.new { 1 }; f.resume; f.resume; rescue FiberError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_fiber_yield_main,
    "begin; Fiber.yield; rescue FiberError; puts 'err'; end",
    "err"
); // can't yield from main fiber
ruby_test!(
    test_fiber_transfer,
    "require 'fiber'; f1 = nil; f2 = Fiber.new { f1.transfer }; f1 = Fiber.new { f2.transfer; 42 }; puts f1.transfer",
    "42"
);
