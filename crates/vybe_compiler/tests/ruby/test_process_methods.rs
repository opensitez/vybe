use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_process_pid, "puts Process.pid.class.name", "Integer");
ruby_test!(test_process_ppid, "puts Process.ppid.class.name", "Integer");
ruby_test!(test_process_uid, "puts Process.uid.class.name", "Integer");
ruby_test!(test_process_gid, "puts Process.gid.class.name", "Integer");
ruby_test!(test_process_euid, "puts Process.euid.class.name", "Integer");
ruby_test!(test_process_egid, "puts Process.egid.class.name", "Integer");
ruby_test!(test_process_clock_gettime, "puts Process.clock_gettime(Process::CLOCK_MONOTONIC).class.name", "Float");
ruby_test!(test_process_times, "puts Process.times.class.name", "Process::Tms");
ruby_test!(test_process_wait2, "pid = fork { exit 42 }; _, status = Process.wait2(pid); puts status.exitstatus", "42");
ruby_test!(test_process_kill, "pid = fork { sleep 10 }; Process.kill('TERM', pid); puts Process.wait(pid)", "pid"); // wait returns pid, but exit status would be term signal. just asserting wait finishes.
ruby_test!(test_process_spawn, "pid = Process.spawn('echo hello > /dev/null'); puts Process.wait(pid)", "pid");
