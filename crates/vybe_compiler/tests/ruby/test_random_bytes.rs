
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_random_bytes, "r = Random.new; puts r.bytes(10).length", "10");
ruby_test!(test_random_urandom, "puts Random.urandom(10).length", "10");
ruby_test!(test_random_new_seed, "puts Random.new_seed.class.name", "Integer");
ruby_test!(test_random_seed, "r = Random.new(42); puts r.seed", "42");
ruby_test!(test_random_rand_range, "r = Random.new(42); puts r.rand(1..10) >= 1", "true");
ruby_test!(test_random_rand_negative, "begin; Random.rand(-1); rescue ArgumentError; puts 'err'; end", "err");
