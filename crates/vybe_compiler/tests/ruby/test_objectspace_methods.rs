use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_objectspace_each_object, "acc = 0; ObjectSpace.each_object(String) { acc += 1 }; puts acc > 0", "true");
ruby_test!(test_objectspace_each_object_no_block, "puts ObjectSpace.each_object(String).class.name", "Enumerator");
ruby_test!(test_objectspace_garbage_collect, "ObjectSpace.garbage_collect; puts 'ok'", "ok");
ruby_test!(test_objectspace_count_objects, "puts ObjectSpace.count_objects.class.name", "Hash");
ruby_test!(test_objectspace_count_objects_key, "puts ObjectSpace.count_objects.key?(:TOTAL)", "true");
ruby_test!(test_objectspace_memsize_of, "require 'objspace'; puts ObjectSpace.memsize_of('hello').class.name", "Integer");
ruby_test!(test_objectspace_reachable_objects_from, "require 'objspace'; a = [1, 2]; puts ObjectSpace.reachable_objects_from(a).class.name", "Array");
