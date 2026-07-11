
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_proc_eql_basic, "p1 = Proc.new { }; p2 = p1.dup; puts p1.eql?(p2)", "false"); // Procs are only equal to themselves
ruby_test!(test_proc_eql_same, "p1 = Proc.new { }; p2 = p1; puts p1.eql?(p2)", "true");
ruby_test!(test_proc_hash_diff, "p1 = Proc.new { }; p2 = p1.dup; puts p1.hash == p2.hash", "false");
ruby_test!(test_method_eql_basic, "class A; def foo; end; end; a = A.new; m1 = a.method(:foo); m2 = a.method(:foo); puts m1 == m2", "true"); // Methods on same receiver/method are equal
ruby_test!(test_method_eql_diff_receiver, "class A; def foo; end; end; m1 = A.new.method(:foo); m2 = A.new.method(:foo); puts m1 == m2", "false");
ruby_test!(test_method_hash_equal, "class A; def foo; end; end; a = A.new; m1 = a.method(:foo); m2 = a.method(:foo); puts m1.hash == m2.hash", "true");
