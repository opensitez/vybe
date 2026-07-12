macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_process_pid, "puts Process.pid.is_a?(Integer)", "true");
ruby_test!(
    test_process_ppid,
    "puts Process.ppid.is_a?(Integer)",
    "true"
);
ruby_test!(test_global_pid, "puts $$.is_a?(Integer)", "true"); // $$ is alias for Process.pid
ruby_test!(
    test_process_clock_gettime,
    "puts Process.clock_gettime(Process::CLOCK_MONOTONIC).is_a?(Float)",
    "true"
);
ruby_test!(
    test_process_times,
    "puts Process.times.is_a?(Process::Tms)",
    "true"
);
