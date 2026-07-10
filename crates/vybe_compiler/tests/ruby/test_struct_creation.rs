use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_struct_creation_basic, "S = Struct.new(:a, :b); puts S.new(1, 2).a", "1");
ruby_test!(test_struct_creation_keyword, "S = Struct.new(:a, :b, keyword_init: true); puts S.new(a: 1, b: 2).b", "2");
ruby_test!(test_struct_creation_block, "S = Struct.new(:a) { def foo; a * 2; end }; puts S.new(3).foo", "6");
ruby_test!(test_struct_creation_anonymous, "s = Struct.new(:a).new(1); puts s.a", "1");
ruby_test!(test_struct_creation_no_args, "S = Struct.new(:a); puts S.new.a.nil?", "true");
ruby_test!(test_struct_members, "S = Struct.new(:a, :b); puts S.members.join('-')", "a-b");
ruby_test!(test_struct_instance_members, "S = Struct.new(:a, :b); puts S.new.members.join('-')", "a-b");
ruby_test!(test_struct_size, "S = Struct.new(:a, :b); puts S.new.size", "2");
ruby_test!(test_struct_length, "S = Struct.new(:a, :b); puts S.new.length", "2");
