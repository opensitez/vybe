macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_hash_iteration_each,
    "acc = []; {a: 1, b: 2}.each { |k, v| acc << \"#{k}:#{v}\" }; puts acc.join('-')",
    "a:1-b:2"
);
ruby_test!(
    test_hash_iteration_each_pair,
    "acc = []; {a: 1, b: 2}.each_pair { |k, v| acc << \"#{k}:#{v}\" }; puts acc.join('-')",
    "a:1-b:2"
);
ruby_test!(
    test_hash_iteration_each_key,
    "acc = []; {a: 1, b: 2}.each_key { |k| acc << k }; puts acc.join('-')",
    "a-b"
);
ruby_test!(
    test_hash_iteration_each_value,
    "acc = []; {a: 1, b: 2}.each_value { |v| acc << v }; puts acc.join('-')",
    "1-2"
);
ruby_test!(
    test_hash_iteration_each_enumerator,
    "puts {a: 1}.each.class.name",
    "Enumerator"
);
ruby_test!(
    test_hash_iteration_each_key_enumerator,
    "puts {a: 1}.each_key.class.name",
    "Enumerator"
);
ruby_test!(
    test_hash_iteration_each_value_enumerator,
    "puts {a: 1}.each_value.class.name",
    "Enumerator"
);
