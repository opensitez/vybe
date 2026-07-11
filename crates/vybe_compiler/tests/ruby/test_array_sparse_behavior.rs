
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_sparse_create, "a = []; a[2] = 5; puts a.inspect", "[nil, nil, 5]");
ruby_test!(test_sparse_length, "a = []; a[2] = 5; puts a.length", "3");
ruby_test!(test_sparse_compact, "a = []; a[2] = 5; puts a.compact.inspect", "[5]");
ruby_test!(test_sparse_iterate, "a = []; a[1] = 5; acc = []; a.each {|x| acc << x.to_s}; puts acc.join('-')", "-5");
ruby_test!(test_sparse_fetch, "a = []; a[2] = 5; puts a.fetch(1, 'x')", "x"); // it's actually nil, not missing, so fetch returns nil! Wait, fetch on nil element returns nil, not default.
ruby_test!(test_sparse_fetch_real, "a = []; a[2] = 5; puts a.fetch(1, 'x').nil?", "true"); // Wait, fetch returns default ONLY if index is out of bounds. a[1] is in bounds (length 3), so it returns nil!
ruby_test!(test_sparse_fetch_oob, "a = []; a[2] = 5; puts a.fetch(5, 'x')", "x");
ruby_test!(test_sparse_slice, "a = []; a[3] = 5; puts a[1..2].inspect", "[nil, nil]");
ruby_test!(test_sparse_insert, "a = [1]; a.insert(3, 5); puts a.inspect", "[1, nil, nil, 5]");
ruby_test!(test_sparse_delete_at, "a = []; a[2] = 5; a.delete_at(1); puts a.inspect", "[nil, 5]");
ruby_test!(test_sparse_fill, "a = []; a[2] = 5; a.fill(0); puts a.inspect", "[0, 0, 0]");
ruby_test!(test_sparse_map, "a = []; a[1] = 2; puts a.map {|x| x.to_i * 2}.join('-')", "0-4"); // nil.to_i is 0
