macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_concat_basic, "s = 'a'; s.concat('b'); puts s", "ab");
ruby_test!(
    test_concat_multiple,
    "s = 'a'; s.concat('b', 'c'); puts s",
    "abc"
);
ruby_test!(test_concat_integer, "s = 'a'; s.concat(98); puts s", "ab"); // 98 is 'b'
ruby_test!(
    test_concat_returns_self,
    "s = 'a'; puts s.concat('b').object_id == s.object_id",
    "true"
);
ruby_test!(test_shovel_basic, "s = 'a'; s << 'b'; puts s", "ab");
ruby_test!(test_shovel_integer, "s = 'a'; s << 98; puts s", "ab");
ruby_test!(test_shovel_chain, "s = 'a'; s << 'b' << 'c'; puts s", "abc");
ruby_test!(test_prepend_basic, "s = 'b'; s.prepend('a'); puts s", "ab");
ruby_test!(
    test_prepend_multiple,
    "s = 'd'; s.prepend('a', 'b', 'c'); puts s",
    "abcd"
);
ruby_test!(
    test_prepend_returns_self,
    "s = 'b'; puts s.prepend('a').object_id == s.object_id",
    "true"
);
ruby_test!(
    test_concat_encoding,
    "s1 = 'a'.force_encoding('UTF-8'); s2 = 'b'.force_encoding('US-ASCII'); puts s1.concat(s2).encoding.name",
    "UTF-8"
);
ruby_test!(
    test_prepend_encoding,
    "s1 = 'b'.force_encoding('UTF-8'); s2 = 'a'.force_encoding('US-ASCII'); puts s1.prepend(s2).encoding.name",
    "UTF-8"
);
