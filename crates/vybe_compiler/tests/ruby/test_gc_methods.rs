use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_gc_start, "GC.start; puts 'ok'", "ok");
ruby_test!(test_gc_enable, "GC.enable; puts 'ok'", "ok");
ruby_test!(test_gc_disable, "puts GC.disable.class.name", "TrueClass"); // returns previous state
ruby_test!(test_gc_stat, "puts GC.stat.class.name", "Hash");
ruby_test!(test_gc_stat_key, "puts GC.stat(:count).class.name", "Integer");
ruby_test!(test_gc_count, "puts GC.count.class.name", "Integer");
ruby_test!(test_gc_latest_gc_info, "puts GC.latest_gc_info.class.name", "Hash");
ruby_test!(test_gc_latest_gc_info_key, "puts GC.latest_gc_info(:major_by).class.name", "Symbol"); // nil or symbol
