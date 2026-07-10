use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_time_edge_at, "t = Time.at(0); puts t.to_i", "0");
ruby_test!(test_time_edge_now, "puts Time.now.class.name", "Time");
ruby_test!(test_time_edge_new_no_args, "puts Time.new.class.name", "Time");
ruby_test!(test_time_edge_utc, "t = Time.utc(2000, 1, 1); puts t.utc?", "true");
ruby_test!(test_time_edge_gm, "t = Time.gm(2000, 1, 1); puts t.utc?", "true");
ruby_test!(test_time_edge_local, "t = Time.local(2000, 1, 1); puts t.class.name", "Time");
ruby_test!(test_time_edge_mktime, "t = Time.mktime(2000, 1, 1); puts t.class.name", "Time");
ruby_test!(test_time_edge_hash, "t = Time.at(0); puts t.hash == Time.at(0).hash", "true");
ruby_test!(test_time_edge_eql, "puts Time.at(0).eql?(Time.at(0))", "true");
ruby_test!(test_time_edge_to_a, "t = Time.utc(2000, 1, 1); a = t.to_a; puts a[0] == 0 && a[1] == 0 && a[2] == 0 && a[3] == 1 && a[4] == 1 && a[5] == 2000", "true");
