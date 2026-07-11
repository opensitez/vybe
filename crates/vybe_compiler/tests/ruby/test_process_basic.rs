
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_process_pid, "puts Process.pid > 0", "true");
ruby_test!(test_process_ppid, "puts Process.ppid > 0", "true");
ruby_test!(test_process_uid, "puts Process.uid >= 0", "true");
ruby_test!(test_process_euid, "puts Process.euid >= 0", "true");
ruby_test!(test_process_gid, "puts Process.gid >= 0", "true");
ruby_test!(test_process_egid, "puts Process.egid >= 0", "true");
ruby_test!(test_process_groups, "puts Process.groups.class.name", "Array");
ruby_test!(test_process_times, "puts Process.times.class.name", "Process::Tms");
ruby_test!(test_process_clock_gettime, "puts Process.clock_gettime(Process::CLOCK_MONOTONIC).class.name", "Float");
ruby_test!(test_process_clock_getres, "puts Process.clock_getres(Process::CLOCK_MONOTONIC).class.name", "Float");
