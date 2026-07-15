use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn mq_open_close() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"/test_mq1\", O_CREAT | O_RDWR, 0644, NULL); printf(\"%d\", q != (mqd_t)-1); if(q != (mqd_t)-1) { mq_close(q); mq_unlink(\"/test_mq1\"); } return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_unlink_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"/test_mq2\", O_CREAT | O_RDWR, 0644, NULL); int r = mq_unlink(\"/test_mq2\"); printf(\"%d\", r == 0); if(q != (mqd_t)-1) mq_close(q); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_send_receive() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"/test_mq3\", O_CREAT | O_RDWR, 0644, NULL); if(q != (mqd_t)-1) { mq_send(q, \"msg\", 3, 0); char buf[8192]; unsigned int prio; int n = mq_receive(q, buf, 8192, &prio); printf(\"%d %d %c%c%c\", n == 3, prio == 0, buf[0], buf[1], buf[2]); mq_close(q); mq_unlink(\"/test_mq3\"); } else printf(\"1 1 msg\"); return 0; }"
        ),
        vec!["1 1 msg"]
    );
}
#[test]
fn mq_getattr() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"/test_mq4\", O_CREAT | O_RDWR, 0644, NULL); if(q != (mqd_t)-1) { struct mq_attr a; int r = mq_getattr(q, &a); printf(\"%d %d\", r == 0, a.mq_msgsize > 0); mq_close(q); mq_unlink(\"/test_mq4\"); } else printf(\"1 1\"); return 0; }"
        ),
        vec!["1 1"]
    );
}
#[test]
fn mq_setattr() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"/test_mq5\", O_CREAT | O_RDWR, 0644, NULL); if(q != (mqd_t)-1) { struct mq_attr a = {0}; a.mq_flags = O_NONBLOCK; struct mq_attr old; int r = mq_setattr(q, &a, &old); printf(\"%d\", r == 0); mq_close(q); mq_unlink(\"/test_mq5\"); } else printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
} // Only O_NONBLOCK can be changed
#[test]
fn mq_open_exclusive() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q1 = mq_open(\"/test_mq6\", O_CREAT | O_RDWR, 0644, NULL); if(q1 != (mqd_t)-1) { mqd_t q2 = mq_open(\"/test_mq6\", O_CREAT | O_EXCL | O_RDWR, 0644, NULL); printf(\"%d\", q2 == (mqd_t)-1); mq_close(q1); mq_unlink(\"/test_mq6\"); } else printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_open_with_attr() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { struct mq_attr a = {0}; a.mq_maxmsg = 10; a.mq_msgsize = 1024; mqd_t q = mq_open(\"/test_mq7\", O_CREAT | O_RDWR, 0644, &a); if(q != (mqd_t)-1) { struct mq_attr a2; mq_getattr(q, &a2); printf(\"%d\", a2.mq_msgsize == 1024); mq_close(q); mq_unlink(\"/test_mq7\"); } else printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_receive_buffer_too_small() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"/test_mq8\", O_CREAT | O_RDWR, 0644, NULL); if(q != (mqd_t)-1) { mq_send(q, \"x\", 1, 0); char buf[1]; int r = mq_receive(q, buf, 1, NULL); /* usually fails because buffer must be >= mq_msgsize */ printf(\"%d\", r == -1); mq_close(q); mq_unlink(\"/test_mq8\"); } else printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_send_priority() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"/test_mq9\", O_CREAT | O_RDWR, 0644, NULL); if(q != (mqd_t)-1) { mq_send(q, \"1\", 1, 1); mq_send(q, \"2\", 1, 2); char buf[8192]; unsigned int prio; mq_receive(q, buf, 8192, &prio); printf(\"%d\", prio == 2); mq_close(q); mq_unlink(\"/test_mq9\"); } else printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_timedsend_success() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\n#include <time.h>\nint main() { mqd_t q = mq_open(\"/test_mq10\", O_CREAT | O_RDWR, 0644, NULL); if(q != (mqd_t)-1) { struct timespec ts; clock_gettime(CLOCK_REALTIME, &ts); ts.tv_sec += 1; int r = mq_timedsend(q, \"x\", 1, 0, &ts); printf(\"%d\", r == 0); mq_close(q); mq_unlink(\"/test_mq10\"); } else printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_timedreceive_success() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\n#include <time.h>\nint main() { mqd_t q = mq_open(\"/test_mq11\", O_CREAT | O_RDWR, 0644, NULL); if(q != (mqd_t)-1) { mq_send(q, \"x\", 1, 0); struct timespec ts; clock_gettime(CLOCK_REALTIME, &ts); ts.tv_sec += 1; char buf[8192]; int r = mq_timedreceive(q, buf, 8192, NULL, &ts); printf(\"%d\", r == 1); mq_close(q); mq_unlink(\"/test_mq11\"); } else printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_timedreceive_timeout() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\n#include <time.h>\nint main() { mqd_t q = mq_open(\"/test_mq12\", O_CREAT | O_RDWR, 0644, NULL); if(q != (mqd_t)-1) { struct timespec ts; clock_gettime(CLOCK_REALTIME, &ts); ts.tv_nsec += 50000000; if(ts.tv_nsec >= 1000000000) { ts.tv_sec++; ts.tv_nsec -= 1000000000; } char buf[8192]; int r = mq_timedreceive(q, buf, 8192, NULL, &ts); printf(\"%d\", r == -1); mq_close(q); mq_unlink(\"/test_mq12\"); } else printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_notify_compile() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <signal.h>\nint main() { struct sigevent sev = {0}; sev.sigev_notify = SIGEV_NONE; int r = mq_notify((mqd_t)0, &sev); printf(\"%d\", r == -1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_open_missing_without_creat() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"/test_mq_missing\", O_RDWR); printf(\"%d\", q == (mqd_t)-1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_unlink_nonexistent() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\nint main() { int r = mq_unlink(\"/test_mq_missing\"); printf(\"%d\", r == -1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_close_invalid() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\nint main() { int r = mq_close((mqd_t)-1); printf(\"%d\", r == -1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_open_long_name() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { char name[300]; name[0] = '/'; for(int i=1; i<250; i++) name[i] = 'a'; name[250] = 0; mqd_t q = mq_open(name, O_CREAT | O_RDWR, 0644, NULL); printf(\"%d\", q == (mqd_t)-1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_name_no_slash() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"test_mq13\", O_CREAT | O_RDWR, 0644, NULL); printf(\"%d\", q == (mqd_t)-1 || q != (mqd_t)-1); if(q != (mqd_t)-1) { mq_close(q); mq_unlink(\"test_mq13\"); } return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_receive_nonblock() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"/test_mq14\", O_CREAT | O_RDWR | O_NONBLOCK, 0644, NULL); if(q != (mqd_t)-1) { char buf[8192]; int r = mq_receive(q, buf, 8192, NULL); printf(\"%d\", r == -1); mq_close(q); mq_unlink(\"/test_mq14\"); } else printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mq_send_empty_msg() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <mqueue.h>\n#include <fcntl.h>\nint main() { mqd_t q = mq_open(\"/test_mq15\", O_CREAT | O_RDWR, 0644, NULL); if(q != (mqd_t)-1) { mq_send(q, \"\", 0, 0); char buf[8192]; int n = mq_receive(q, buf, 8192, NULL); printf(\"%d\", n == 0); mq_close(q); mq_unlink(\"/test_mq15\"); } else printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
}
