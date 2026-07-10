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
ruby_test!(test_gc_enable_disable, "GC.disable; GC.enable; puts 'ok'", "ok");
ruby_test!(test_gc_stat, "puts GC.stat.is_a?(Hash)", "true");
ruby_test!(test_gc_latest_gc_info, "puts GC.latest_gc_info.is_a?(Hash) || GC.latest_gc_info.nil?", "true"); // might be nil if no GC happened
ruby_test!(test_gc_count, "puts GC.count.is_a?(Integer)", "true");
