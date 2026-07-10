use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_exception_classes_standarderror, "begin; raise StandardError; rescue => e; puts e.class.name; end", "StandardError");
ruby_test!(test_exception_classes_argumenterror, "begin; raise ArgumentError; rescue => e; puts e.class.name; end", "ArgumentError");
ruby_test!(test_exception_classes_typeerror, "begin; raise TypeError; rescue => e; puts e.class.name; end", "TypeError");
ruby_test!(test_exception_classes_runtimeerror, "begin; raise 'err'; rescue => e; puts e.class.name; end", "RuntimeError"); // raise with string creates RuntimeError
ruby_test!(test_exception_classes_nomethoderror, "begin; Object.new.does_not_exist; rescue => e; puts e.class.name; end", "NoMethodError");
ruby_test!(test_exception_classes_nameerror, "begin; does_not_exist; rescue => e; puts e.class.name; end", "NameError");
ruby_test!(test_exception_classes_zerodivisionerror, "begin; 1 / 0; rescue => e; puts e.class.name; end", "ZeroDivisionError");
ruby_test!(test_exception_classes_indexerror, "begin; [].fetch(1); rescue => e; puts e.class.name; end", "IndexError");
ruby_test!(test_exception_classes_keyerror, "begin; {}.fetch(:a); rescue => e; puts e.class.name; end", "KeyError");
ruby_test!(test_exception_classes_rescue_default, "begin; raise Exception; rescue; puts 'caught'; else; puts 'not caught'; end", "not caught"); // bare rescue only catches StandardError
