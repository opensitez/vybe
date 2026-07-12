macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_struct_to_h,
    "S = Struct.new(:a, :b); puts S.new(1, 2).to_h.map { |k, v| \"#{k}:#{v}\" }.join('-')",
    "a:1-b:2"
);
ruby_test!(
    test_struct_to_a,
    "S = Struct.new(:a, :b); puts S.new(1, 2).to_a.join('-')",
    "1-2"
);
ruby_test!(
    test_struct_members,
    "S = Struct.new(:a, :b); puts S.members.join('-')",
    "a-b"
);
ruby_test!(
    test_struct_members_instance,
    "S = Struct.new(:a, :b); puts S.new.members.join('-')",
    "a-b"
);
ruby_test!(
    test_struct_select,
    "S = Struct.new(:a, :b, :c); puts S.new(1, 2, 3).select { |v| v > 1 }.join('-')",
    "2-3"
);
ruby_test!(
    test_struct_size,
    "S = Struct.new(:a, :b); puts S.new.size",
    "2"
);
ruby_test!(
    test_struct_length,
    "S = Struct.new(:a, :b); puts S.new.length",
    "2"
);
