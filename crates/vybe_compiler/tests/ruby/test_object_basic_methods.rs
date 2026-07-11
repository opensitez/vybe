
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_object_id_basic, "puts Object.new.object_id.is_a?(Integer)", "true");
ruby_test!(test_object_id_same, "o = Object.new; puts o.object_id == o.object_id", "true");
ruby_test!(test_object_id_diff, "puts Object.new.object_id == Object.new.object_id", "false");
ruby_test!(test_object_class, "puts Object.new.class.name", "Object");
ruby_test!(test_object_nil, "puts Object.new.nil?", "false");
ruby_test!(test_object_is_a, "puts Object.new.is_a?(Object)", "true");
ruby_test!(test_object_is_a_false, "puts Object.new.is_a?(String)", "false");
ruby_test!(test_object_kind_of, "puts Object.new.kind_of?(Object)", "true");
ruby_test!(test_object_instance_of, "puts Object.new.instance_of?(Object)", "true");
ruby_test!(test_object_instance_of_subclass, "class A; end; puts A.new.instance_of?(Object)", "false");
ruby_test!(test_object_tap, "puts Object.new.tap { |o| o }.class.name", "Object");
