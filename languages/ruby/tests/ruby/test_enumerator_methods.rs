macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_enumerator_yielder,
    "e = Enumerator.new { |y| y << 1; y << 2 }; puts e.to_a.join('-')",
    "1-2"
);
ruby_test!(
    test_enumerator_yield,
    "e = Enumerator.new { |y| y.yield 1; y.yield 2 }; puts e.to_a.join('-')",
    "1-2"
);
ruby_test!(
    test_enumerator_next,
    "e = [1, 2].each; puts e.next; puts e.next",
    "1\n2"
);
ruby_test!(
    test_enumerator_rewind,
    "e = [1, 2].each; e.next; e.rewind; puts e.next",
    "1"
);
ruby_test!(
    test_enumerator_peek,
    "e = [1, 2].each; puts e.peek; puts e.next",
    "1\n1"
);
ruby_test!(
    test_enumerator_next_stopiteration,
    "e = [1].each; e.next; begin; e.next; rescue StopIteration; puts 'err'; end",
    "err"
);
ruby_test!(
    test_enumerator_peek_stopiteration,
    "e = [1].each; e.next; begin; e.peek; rescue StopIteration; puts 'err'; end",
    "err"
);
ruby_test!(
    test_enumerator_with_index,
    "puts [1, 2].each.with_index { |x, i| \"#{x}:#{i}\" }.join('-')",
    "1:0-2:1"
); // each without block returns enum, but then with_index needs to map? wait with_index without block returns enum. with block it iterates. but array.each.with_index doesn't return array, it returns what each returns (array). so map is better here.
ruby_test!(
    test_enumerator_with_index_map,
    "puts [1, 2].map.with_index { |x, i| \"#{x}:#{i}\" }.join('-')",
    "1:0-2:1"
);
ruby_test!(
    test_enumerator_with_object,
    "puts [1, 2].each.with_object([]) { |x, arr| arr << x * 2 }.join('-')",
    "2-4"
);
ruby_test!(test_enumerator_size, "puts [1, 2].each.size", "2");
ruby_test!(
    test_enumerator_chain,
    "puts [1, 2].each.chain([3, 4]).to_a.join('-')",
    "1-2-3-4"
);
ruby_test!(
    test_enumerator_chain_plus,
    "puts ([1, 2].each + [3, 4].each).to_a.join('-')",
    "1-2-3-4"
);
