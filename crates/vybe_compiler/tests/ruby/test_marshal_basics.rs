macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_marshal_dump_load_string,
    "d = Marshal.dump('hello'); puts Marshal.load(d)",
    "hello"
);
ruby_test!(
    test_marshal_dump_load_integer,
    "d = Marshal.dump(123); puts Marshal.load(d)",
    "123"
);
ruby_test!(
    test_marshal_dump_load_array,
    "d = Marshal.dump([1, 'a', :b]); puts Marshal.load(d).join('-')",
    "1-a-b"
);
ruby_test!(
    test_marshal_dump_load_hash,
    "d = Marshal.dump({a: 1}); puts Marshal.load(d)[:a]",
    "1"
);
ruby_test!(
    test_marshal_dump_load_object,
    "class A; attr_accessor :x; end; a = A.new; a.x = 1; d = Marshal.dump(a); a2 = Marshal.load(d); puts a2.x",
    "1"
);
ruby_test!(
    test_marshal_dump_unsupported,
    "begin; Marshal.dump(Proc.new {}); rescue TypeError; puts 'err'; end",
    "err"
);
