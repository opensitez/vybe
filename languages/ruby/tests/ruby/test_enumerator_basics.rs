macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_enumerator_basic,
    "e = Enumerator.new { |y| y << 1; y << 2; y << 3 }; puts e.to_a.join('-')",
    "1-2-3"
);
ruby_test!(
    test_enumerator_next,
    "e = [1, 2].to_enum; puts \"#{e.next}-#{e.next}\"",
    "1-2"
);
ruby_test!(
    test_enumerator_rewind,
    "e = [1].to_enum; e.next; e.rewind; puts e.next",
    "1"
);
ruby_test!(
    test_enumerator_peek,
    "e = [1, 2].to_enum; puts \"#{e.peek}-#{e.next}-#{e.peek}\"",
    "1-1-2"
);
ruby_test!(
    test_enumerator_with_index,
    "e = [10, 20].to_enum; acc = []; e.with_index { |v, i| acc << \"#{v}:#{i}\" }; puts acc.join('-')",
    "10:0-20:1"
);
ruby_test!(
    test_enumerator_size,
    "e = [1, 2, 3].to_enum; puts e.size",
    "3"
);
