use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_unpack_c, "puts 'abc'.unpack('C*').join(',')", "97,98,99");
ruby_test!(test_unpack_c_signed, "puts \"\\xFF\".unpack('c').join(',')", "-1");
ruby_test!(test_unpack_hex, "puts 'abc'.unpack('H*').first", "616263");
ruby_test!(test_unpack_hex_nibble, "puts 'abc'.unpack('h*').first", "162636");
ruby_test!(test_unpack_base64, "puts 'hello'.unpack('m0').first", "aGVsbG8=");
ruby_test!(test_unpack_pack_roundtrip, "puts ['abc'.unpack('H*').first].pack('H*')", "abc");
ruby_test!(test_unpack_n, "puts \"\\x01\\x02\".unpack('n').first", "258"); // Network byte order
ruby_test!(test_unpack_v, "puts \"\\x01\\x02\".unpack('v').first", "513"); // Little endian
ruby_test!(test_unpack_a, "puts 'hello  '.unpack('A5').first", "hello");
ruby_test!(test_unpack_z, "puts \"hello\\x00world\".unpack('Z*').first", "hello");
ruby_test!(test_unpack_x, "puts 'abcdef'.unpack('x2 C').first", "99"); // Skip 2 bytes, read 'c'
ruby_test!(test_unpack_multiple, "puts 'abc'.unpack('C C C').join('-')", "97-98-99");
