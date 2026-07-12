macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_process_status_to_i,
    "system('exit 42'); puts $?.to_i >> 8",
    "42"
);
ruby_test!(
    test_process_status_exitstatus,
    "system('exit 42'); puts $?.exitstatus",
    "42"
);
ruby_test!(
    test_process_status_success,
    "system('exit 0'); puts $?.success?",
    "true"
);
ruby_test!(
    test_process_status_success_false,
    "system('exit 1'); puts $?.success?",
    "false"
);
ruby_test!(
    test_process_status_pid,
    "system('exit 0'); puts $?.pid > 0",
    "true"
);
ruby_test!(
    test_process_status_exited,
    "system('exit 0'); puts $?.exited?",
    "true"
);
ruby_test!(
    test_process_status_signaled,
    "system('exit 0'); puts $?.signaled?",
    "false"
);
ruby_test!(
    test_process_status_stopped,
    "system('exit 0'); puts $?.stopped?",
    "false"
);
ruby_test!(
    test_process_status_termsig,
    "system('exit 0'); puts $?.termsig.nil?",
    "true"
);
ruby_test!(
    test_process_status_stopsig,
    "system('exit 0'); puts $?.stopsig.nil?",
    "true"
);
