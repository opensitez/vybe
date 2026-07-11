
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_struct_iteration_each, "S = Struct.new(:a, :b); acc = []; S.new(1, 2).each { |v| acc << v }; puts acc.join('-')", "1-2");
ruby_test!(test_struct_iteration_each_pair, "S = Struct.new(:a, :b); acc = []; S.new(1, 2).each_pair { |k, v| acc << \"#{k}:#{v}\" }; puts acc.join('-')", "a:1-b:2");
ruby_test!(test_struct_iteration_select, "S = Struct.new(:a, :b, :c); puts S.new(1, 2, 3).select { |v| v > 1 }.join('-')", "2-3");
ruby_test!(test_struct_iteration_to_a, "S = Struct.new(:a, :b); puts S.new(1, 2).to_a.join('-')", "1-2");
ruby_test!(test_struct_iteration_values, "S = Struct.new(:a, :b); puts S.new(1, 2).values.join('-')", "1-2");
ruby_test!(test_struct_iteration_values_at, "S = Struct.new(:a, :b, :c); puts S.new(1, 2, 3).values_at(0, 2).join('-')", "1-3");
ruby_test!(test_struct_iteration_to_h, "S = Struct.new(:a, :b); h = S.new(1, 2).to_h; puts \"#{h[:a]}-#{h[:b]}\"", "1-2");
ruby_test!(test_struct_iteration_to_h_block, "S = Struct.new(:a, :b); h = S.new(1, 2).to_h { |k, v| [k.to_s, v * 2] }; puts \"#{h['a']}-#{h['b']}\"", "2-4");
