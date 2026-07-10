use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_rand_basic, "puts rand >= 0.0 && rand < 1.0", "true");
ruby_test!(test_rand_integer, "puts [0, 1, 2].include?(rand(3))", "true");
ruby_test!(test_rand_range, "puts (10..20).include?(rand(10..20))", "true");
ruby_test!(test_srand_reproducibility, "srand(123); a = rand; srand(123); b = rand; puts a == b", "true");
ruby_test!(test_random_class_rand, "r = Random.new(123); a = r.rand; r2 = Random.new(123); b = r2.rand; puts a == b", "true");
ruby_test!(test_random_bytes, "puts Random.new.bytes(5).bytesize", "5");
