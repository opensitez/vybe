
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_random_rand_float, "r = Random.new(42); f = r.rand; puts f >= 0.0 && f < 1.0", "true");
ruby_test!(test_random_rand_int, "r = Random.new(42); i = r.rand(10); puts i >= 0 && i < 10", "true");
ruby_test!(test_random_rand_range, "r = Random.new(42); i = r.rand(10..20); puts i >= 10 && i <= 20", "true");
ruby_test!(test_random_bytes, "r = Random.new(42); b = r.bytes(5); puts b.length", "5");
ruby_test!(test_random_seed, "r = Random.new(42); puts r.seed", "42");
ruby_test!(test_random_class_rand_float, "f = Random.rand; puts f >= 0.0 && f < 1.0", "true");
ruby_test!(test_random_class_rand_int, "i = Random.rand(10); puts i >= 0 && i < 10", "true");
ruby_test!(test_random_class_bytes, "b = Random.bytes(5); puts b.length", "5");
ruby_test!(test_random_class_new_seed, "s = Random.new_seed; puts s.class.name", "Integer");
ruby_test!(test_random_srand, "old = srand(42); new_old = srand(42); puts new_old == 42", "true");
ruby_test!(test_kernel_rand, "old = srand(42); v1 = rand(10); srand(42); v2 = rand(10); puts v1 == v2", "true");
