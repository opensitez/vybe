use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_fiber_basic, "f = Fiber.new { |x| x * 2 }; puts f.resume(3)", "6");
ruby_test!(test_fiber_yield, "f = Fiber.new { Fiber.yield 1; 2 }; puts \"#{f.resume}-#{f.resume}\"", "1-2");
ruby_test!(test_fiber_alive, "f = Fiber.new { Fiber.yield; 2 }; puts f.alive?; f.resume; puts f.alive?; f.resume; puts f.alive?", "true\ntrue\nfalse"); // Actually, run_ruby_one checks final output. Let's do array.
ruby_test!(test_fiber_alive_array, "f = Fiber.new { Fiber.yield; 2 }; acc = [f.alive?]; f.resume; acc << f.alive?; f.resume; acc << f.alive?; puts acc.join('-')", "true-true-false");
ruby_test!(test_fiber_current, "puts Fiber.current.class.name", "Fiber");
ruby_test!(test_fiber_resume_args, "f = Fiber.new { |a| a + Fiber.yield(a*2) }; puts \"#{f.resume(2)}-#{f.resume(3)}\"", "4-5");
