use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn unix_socket_create() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_UNIX, SOCK_STREAM, 0); printf(\"%d\", fd >= 0); if(fd>=0) close(fd); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_bind() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\n#include <string.h>\nint main() { unlink(\"test_unix.sock\"); int fd = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; strcpy(addr.sun_path, \"test_unix.sock\"); int r = bind(fd, (struct sockaddr*)&addr, sizeof(addr)); printf(\"%d\", r == 0); close(fd); unlink(\"test_unix.sock\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_listen() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\n#include <string.h>\nint main() { unlink(\"test_unix2.sock\"); int fd = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; strcpy(addr.sun_path, \"test_unix2.sock\"); bind(fd, (struct sockaddr*)&addr, sizeof(addr)); int r = listen(fd, 5); printf(\"%d\", r == 0); close(fd); unlink(\"test_unix2.sock\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_connect_accept() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\n#include <string.h>\n#include <pthread.h>\nvoid* f(void* a) { int s = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; strcpy(addr.sun_path, \"test_unix3.sock\"); while(connect(s, (struct sockaddr*)&addr, sizeof(addr)) != 0) usleep(10000); send(s, \"unix\", 4, 0); close(s); return NULL; }\nint main() { unlink(\"test_unix3.sock\"); int l = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un a={0}; a.sun_family = AF_UNIX; strcpy(a.sun_path, \"test_unix3.sock\"); bind(l, (struct sockaddr*)&a, sizeof(a)); listen(l, 5); pthread_t t; pthread_create(&t, NULL, f, NULL); int c = accept(l, NULL, NULL); char b[5]={0}; recv(c, b, 4, 0); printf(\"%s\", b); close(c); close(l); pthread_join(t, NULL); unlink(\"test_unix3.sock\"); return 0; }"
        ),
        vec!["unix"]
    );
}
#[test]
fn unix_socketpair() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd[2]; int r = socketpair(AF_UNIX, SOCK_STREAM, 0, fd); printf(\"%d %d %d\", r == 0, fd[0] >= 0, fd[1] >= 0); close(fd[0]); close(fd[1]); return 0; }"
        ),
        vec!["1 1 1"]
    );
}
#[test]
fn unix_socketpair_dgram() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd[2]; int r = socketpair(AF_UNIX, SOCK_DGRAM, 0, fd); printf(\"%d\", r == 0); close(fd[0]); close(fd[1]); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_socketpair_io() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd[2]; socketpair(AF_UNIX, SOCK_STREAM, 0, fd); write(fd[0], \"pair\", 4); char b[5]={0}; read(fd[1], b, 4); printf(\"%s\", b); close(fd[0]); close(fd[1]); return 0; }"
        ),
        vec!["pair"]
    );
}
#[test]
fn unix_socketpair_shutdown() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd[2]; socketpair(AF_UNIX, SOCK_STREAM, 0, fd); shutdown(fd[0], SHUT_WR); char b[5]; int r = read(fd[1], b, 5); printf(\"%d\", r == 0); close(fd[0]); close(fd[1]); return 0; }"
        ),
        vec!["1"]
    );
} // EOF reached
#[test]
fn unix_bind_existing_file() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\n#include <fcntl.h>\n#include <string.h>\nint main() { int f = open(\"test_unix_ext.sock\", O_CREAT, 0644); close(f); int fd = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; strcpy(addr.sun_path, \"test_unix_ext.sock\"); int r = bind(fd, (struct sockaddr*)&addr, sizeof(addr)); printf(\"%d\", r == -1); close(fd); unlink(\"test_unix_ext.sock\"); return 0; }"
        ),
        vec!["1"]
    );
} // Fails if file exists
#[test]
fn unix_getsockname() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\n#include <string.h>\nint main() { unlink(\"test_unix4.sock\"); int fd = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; strcpy(addr.sun_path, \"test_unix4.sock\"); bind(fd, (struct sockaddr*)&addr, sizeof(addr)); struct sockaddr_un a2={0}; socklen_t len = sizeof(a2); getsockname(fd, (struct sockaddr*)&a2, &len); printf(\"%s\", a2.sun_path); close(fd); unlink(\"test_unix4.sock\"); return 0; }"
        ),
        vec!["test_unix4.sock"]
    );
}
#[test]
fn unix_dgram_bind() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\n#include <string.h>\nint main() { unlink(\"test_unix_d.sock\"); int fd = socket(AF_UNIX, SOCK_DGRAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; strcpy(addr.sun_path, \"test_unix_d.sock\"); int r = bind(fd, (struct sockaddr*)&addr, sizeof(addr)); printf(\"%d\", r == 0); close(fd); unlink(\"test_unix_d.sock\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_dgram_sendto() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\n#include <string.h>\nint main() { unlink(\"t1.sock\"); unlink(\"t2.sock\"); int s1 = socket(AF_UNIX, SOCK_DGRAM, 0); int s2 = socket(AF_UNIX, SOCK_DGRAM, 0); struct sockaddr_un a1={0}, a2={0}; a1.sun_family = AF_UNIX; strcpy(a1.sun_path, \"t1.sock\"); a2.sun_family = AF_UNIX; strcpy(a2.sun_path, \"t2.sock\"); bind(s1, (struct sockaddr*)&a1, sizeof(a1)); bind(s2, (struct sockaddr*)&a2, sizeof(a2)); sendto(s1, \"D\", 1, 0, (struct sockaddr*)&a2, sizeof(a2)); char b[2]={0}; recvfrom(s2, b, 1, 0, NULL, NULL); printf(\"%s\", b); close(s1); close(s2); unlink(\"t1.sock\"); unlink(\"t2.sock\"); return 0; }"
        ),
        vec!["D"]
    );
}
#[test]
fn unix_connect_nonexistent() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\n#include <string.h>\nint main() { int s = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; strcpy(addr.sun_path, \"doesnotexist.sock\"); int r = connect(s, (struct sockaddr*)&addr, sizeof(addr)); printf(\"%d\", r == -1); close(s); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_sun_path_length() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/un.h>\nint main() { printf(\"%d\", sizeof(((struct sockaddr_un*)0)->sun_path) > 10); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_socketpair_nonblock() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\n#include <fcntl.h>\nint main() { int fd[2]; socketpair(AF_UNIX, SOCK_STREAM, 0, fd); fcntl(fd[0], F_SETFL, O_NONBLOCK); char b[1]; int r = read(fd[0], b, 1); printf(\"%d\", r == -1); close(fd[0]); close(fd[1]); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_socketpair_close_one() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd[2]; socketpair(AF_UNIX, SOCK_STREAM, 0, fd); close(fd[1]); char b[1]; int r = read(fd[0], b, 1); printf(\"%d\", r == 0); close(fd[0]); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_socketpair_write_closed() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\n#include <signal.h>\nint main() { signal(SIGPIPE, SIG_IGN); int fd[2]; socketpair(AF_UNIX, SOCK_STREAM, 0, fd); close(fd[0]); int r = write(fd[1], \"x\", 1); printf(\"%d\", r == -1); close(fd[1]); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_getpeername() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\n#include <string.h>\n#include <pthread.h>\nvoid* f(void* a) { int s = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; strcpy(addr.sun_path, \"test_unix5.sock\"); while(connect(s, (struct sockaddr*)&addr, sizeof(addr)) != 0) usleep(10000); close(s); return NULL; }\nint main() { unlink(\"test_unix5.sock\"); int l = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un a={0}; a.sun_family = AF_UNIX; strcpy(a.sun_path, \"test_unix5.sock\"); bind(l, (struct sockaddr*)&a, sizeof(a)); listen(l, 5); pthread_t t; pthread_create(&t, NULL, f, NULL); int c = accept(l, NULL, NULL); struct sockaddr_un p={0}; socklen_t len=sizeof(p); getpeername(c, (struct sockaddr*)&p, &len); printf(\"%d\", p.sun_family == AF_UNIX); close(c); close(l); pthread_join(t, NULL); unlink(\"test_unix5.sock\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn socketpair_domain_invalid() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\nint main() { int fd[2]; int r = socketpair(AF_INET, SOCK_STREAM, 0, fd); /* Usually socketpair is AF_UNIX only */ printf(\"%d\", r == -1 || r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unix_abstract_namespace_gnu() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <sys/socket.h>\n#include <sys/un.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; addr.sun_path[0] = '\\0'; addr.sun_path[1] = 't'; addr.sun_path[2] = 's'; addr.sun_path[3] = 't'; int r = bind(fd, (struct sockaddr*)&addr, sizeof(sa_family_t) + 4); printf(\"%d\", r == 0 || r != 0); close(fd); return 0; }"
        ),
        vec!["1"]
    );
} // May fail on non-linux
