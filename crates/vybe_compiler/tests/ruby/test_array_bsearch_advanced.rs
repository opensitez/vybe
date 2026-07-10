use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_bsearch_find_exact, "puts [1, 3, 5, 7, 9].bsearch {|x| 5 <=> x}", "5");
ruby_test!(test_bsearch_find_less, "puts [1, 3, 5, 7, 9].bsearch {|x| 4 <=> x}.nil?", "true"); // Wait, bsearch block for exact match must return -1, 0, 1.
// If block returns a number, it behaves as Find-Minimum or Find-Exact.
// Wait, `|x| 5 <=> x` means if x is 3, 5 <=> 3 is 1 (go right). If x is 7, 5 <=> 7 is -1 (go left).
// This is exactly what bsearch needs.
ruby_test!(test_bsearch_find_boolean_first_true, "puts [1, 3, 5, 7, 9].bsearch {|x| x >= 4}", "5");
ruby_test!(test_bsearch_boolean_all_false, "puts [1, 3, 5, 7, 9].bsearch {|x| x >= 10}.nil?", "true");
ruby_test!(test_bsearch_boolean_all_true, "puts [1, 3, 5, 7, 9].bsearch {|x| x >= 0}", "1");
ruby_test!(test_bsearch_spaceship_found, "puts [1, 3, 5, 7, 9].bsearch {|x| 7 <=> x}", "7");
ruby_test!(test_bsearch_spaceship_not_found, "puts [1, 3, 5, 7, 9].bsearch {|x| 6 <=> x}.nil?", "true");
ruby_test!(test_bsearch_spaceship_first, "puts [1, 3, 5, 7, 9].bsearch {|x| 1 <=> x}", "1");
ruby_test!(test_bsearch_spaceship_last, "puts [1, 3, 5, 7, 9].bsearch {|x| 9 <=> x}", "9");
ruby_test!(test_bsearch_empty, "puts [].bsearch {|x| x >= 1}.nil?", "true");
ruby_test!(test_bsearch_duplicate_boolean, "puts [1, 5, 5, 5, 9].bsearch {|x| x >= 4}", "5"); // Returns one of the 5s (usually first meeting cond)
ruby_test!(test_bsearch_string, "puts ['a', 'c', 'e'].bsearch {|x| x >= 'b'}", "c");
