
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_struct_basic, "S = Struct.new(:x, :y); puts S.new(1, 2).x", "1");
ruby_test!(test_struct_set, "S = Struct.new(:x, :y); s = S.new; s.x = 1; puts s.x", "1");
ruby_test!(test_struct_members, "S = Struct.new(:x, :y); puts S.new.members.join('-')", "x-y");
ruby_test!(test_struct_values, "S = Struct.new(:x, :y); puts S.new(1, 2).values.join('-')", "1-2");
ruby_test!(test_struct_each, "S = Struct.new(:x, :y); acc = []; S.new(1, 2).each { |v| acc << v }; puts acc.join('-')", "1-2");
ruby_test!(test_struct_each_pair, "S = Struct.new(:x, :y); acc = []; S.new(1, 2).each_pair { |k, v| acc << \"#{k}:#{v}\" }; puts acc.join('-')", "x:1-y:2");
ruby_test!(test_struct_brackets, "S = Struct.new(:x, :y); s = S.new(1, 2); puts s[:y]", "2");
ruby_test!(test_struct_brackets_set, "S = Struct.new(:x, :y); s = S.new; s[:y] = 2; puts s.y", "2");
ruby_test!(test_struct_block, "S = Struct.new(:x) { def foo; x * 2; end }; puts S.new(3).foo", "6");
