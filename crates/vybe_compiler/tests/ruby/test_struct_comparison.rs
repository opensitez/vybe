
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_struct_comparison_eq, "S = Struct.new(:a, :b); puts S.new(1, 2) == S.new(1, 2)", "true");
ruby_test!(test_struct_comparison_not_eq, "S = Struct.new(:a, :b); puts S.new(1, 2) == S.new(2, 1)", "false");
ruby_test!(test_struct_comparison_eql, "S = Struct.new(:a, :b); puts S.new(1, 2).eql?(S.new(1, 2))", "true");
ruby_test!(test_struct_comparison_hash, "S = Struct.new(:a, :b); puts S.new(1, 2).hash == S.new(1, 2).hash", "true");
ruby_test!(test_struct_comparison_type_mismatch, "S1 = Struct.new(:a); S2 = Struct.new(:a); puts S1.new(1) == S2.new(1)", "false");
ruby_test!(test_struct_comparison_eqq, "S = Struct.new(:a); puts S === S.new(1)", "true");
