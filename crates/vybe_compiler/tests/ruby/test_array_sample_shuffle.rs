use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_sample_single, "puts [1, 2, 3].include?([1, 2, 3].sample)", "true");
ruby_test!(test_sample_multiple, "puts [1, 2, 3].sample(2).length", "2");
ruby_test!(test_sample_all, "puts [1, 2, 3].sample(3).length", "3");
ruby_test!(test_sample_over, "puts [1, 2, 3].sample(5).length", "3");
ruby_test!(test_sample_zero, "puts [1, 2, 3].sample(0).length", "0");
ruby_test!(test_sample_empty_single, "puts [].sample.nil?", "true");
ruby_test!(test_sample_empty_multiple, "puts [].sample(2).length", "0");
ruby_test!(test_sample_random_parameter, "puts [1, 2, 3].sample(random: Random.new(1)).nil?", "false"); // Just checks it runs
ruby_test!(test_shuffle_basic, "puts [1, 2, 3].shuffle.length", "3");
ruby_test!(test_shuffle_contains_same, "a = [1, 2, 3]; puts (a.shuffle - a).empty? && (a - a.shuffle).empty?", "true");
ruby_test!(test_shuffle_empty, "puts [].shuffle.length", "0");
ruby_test!(test_shuffle_random_parameter, "puts [1, 2, 3].shuffle(random: Random.new(1)).length", "3");
ruby_test!(test_shuffle_bang_mutates, "a = [1, 2, 3]; a.shuffle!; puts a.length", "3");
ruby_test!(test_shuffle_bang_returns_self, "a = [1, 2, 3]; puts a.shuffle!.object_id == a.object_id", "true");
