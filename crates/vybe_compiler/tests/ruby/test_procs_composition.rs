macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_proc_composition_right,
    "p1 = proc {|x| x * 2 }; p2 = proc {|x| x + 1 }; puts (p1 >> p2).call(3)",
    "7"
); // (3 * 2) + 1 = 7
ruby_test!(
    test_proc_composition_left,
    "p1 = proc {|x| x * 2 }; p2 = proc {|x| x + 1 }; puts (p1 << p2).call(3)",
    "8"
); // (3 + 1) * 2 = 8
ruby_test!(
    test_method_composition_right,
    "class A; def f1(x); x * 2; end; def f2(x); x + 1; end; end; a = A.new; m1 = a.method(:f1); m2 = a.method(:f2); puts (m1 >> m2).call(3)",
    "7"
);
ruby_test!(
    test_method_composition_left,
    "class A; def f1(x); x * 2; end; def f2(x); x + 1; end; end; a = A.new; m1 = a.method(:f1); m2 = a.method(:f2); puts (m1 << m2).call(3)",
    "8"
);
ruby_test!(
    test_proc_method_composition,
    "p = proc {|x| x * 2 }; class A; def f(x); x + 1; end; end; m = A.new.method(:f); puts (p >> m).call(3)",
    "7"
);
