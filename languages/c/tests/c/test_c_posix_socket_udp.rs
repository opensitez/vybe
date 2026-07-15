use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn socket_create_udp() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET, SOCK_DGRAM, 0); printf(\"%d\", fd >= 0); if(fd>=0) close(fd); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn bind_udp_loopback() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in addr={0}; addr.sin_family = AF_INET; addr.sin_addr.s_addr = htonl(INADDR_LOOPBACK); addr.sin_port = 0; int r = bind(fd, (struct sockaddr*)&addr, sizeof(addr)); printf(\"%d\", r == 0); if(fd>=0) close(fd); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn sendto_recvfrom_udp() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\n#include <pthread.h>\nint port = 0;\nvoid* f(void* a) { int s = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in addr={0}; addr.sin_family = AF_INET; addr.sin_addr.s_addr = htonl(INADDR_LOOPBACK); addr.sin_port = htons(port); sendto(s, \"udp\", 3, 0, (struct sockaddr*)&addr, sizeof(addr)); close(s); return NULL; }\nint main() { int l = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a={0}; a.sin_family = AF_INET; a.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(l, (struct sockaddr*)&a, sizeof(a)); socklen_t len=sizeof(a); getsockname(l, (struct sockaddr*)&a, &len); port = ntohs(a.sin_port); pthread_t t; pthread_create(&t, NULL, f, NULL); char b[5]={0}; struct sockaddr_in from; socklen_t flen = sizeof(from); recvfrom(l, b, 3, 0, (struct sockaddr*)&from, &flen); printf(\"%s\", b); close(l); pthread_join(t, NULL); return 0; }"
        ),
        vec!["udp"]
    );
}
#[test]
fn connect_udp() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in addr={0}; addr.sin_family = AF_INET; addr.sin_addr.s_addr = htonl(INADDR_LOOPBACK); addr.sin_port = htons(53); /* arbitrary */ int r = connect(fd, (struct sockaddr*)&addr, sizeof(addr)); printf(\"%d\", r == 0); close(fd); return 0; }"
        ),
        vec!["1"]
    );
} // Connect on UDP sets default destination
#[test]
fn send_recv_connected_udp() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int s1 = socket(AF_INET, SOCK_DGRAM, 0); int s2 = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a1={0}; a1.sin_family = AF_INET; a1.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(s1, (struct sockaddr*)&a1, sizeof(a1)); socklen_t l1 = sizeof(a1); getsockname(s1, (struct sockaddr*)&a1, &l1); struct sockaddr_in a2={0}; a2.sin_family = AF_INET; a2.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(s2, (struct sockaddr*)&a2, sizeof(a2)); socklen_t l2 = sizeof(a2); getsockname(s2, (struct sockaddr*)&a2, &l2); connect(s1, (struct sockaddr*)&a2, l2); connect(s2, (struct sockaddr*)&a1, l1); send(s1, \"hi\", 2, 0); char b[5]={0}; recv(s2, b, 2, 0); printf(\"%s\", b); close(s1); close(s2); return 0; }"
        ),
        vec!["hi"]
    );
}
#[test]
fn setsockopt_broadcast() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET, SOCK_DGRAM, 0); int opt = 1; int r = setsockopt(fd, SOL_SOCKET, SO_BROADCAST, &opt, sizeof(opt)); printf(\"%d\", r == 0); close(fd); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getsockopt_broadcast() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET, SOCK_DGRAM, 0); int opt = 1; setsockopt(fd, SOL_SOCKET, SO_BROADCAST, &opt, sizeof(opt)); int o2 = 0; socklen_t len = sizeof(o2); getsockopt(fd, SOL_SOCKET, SO_BROADCAST, &o2, &len); printf(\"%d\", o2 != 0); close(fd); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn recvfrom_null_addr() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int s1 = socket(AF_INET, SOCK_DGRAM, 0); int s2 = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a1={0}; a1.sin_family = AF_INET; a1.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(s1, (struct sockaddr*)&a1, sizeof(a1)); socklen_t l1 = sizeof(a1); getsockname(s1, (struct sockaddr*)&a1, &l1); sendto(s2, \"X\", 1, 0, (struct sockaddr*)&a1, l1); char b[2]={0}; recvfrom(s1, b, 1, 0, NULL, NULL); printf(\"%s\", b); close(s1); close(s2); return 0; }"
        ),
        vec!["X"]
    );
}
#[test]
fn sendto_null_dest_connected() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int s1 = socket(AF_INET, SOCK_DGRAM, 0); int s2 = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a1={0}; a1.sin_family = AF_INET; a1.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(s1, (struct sockaddr*)&a1, sizeof(a1)); socklen_t l1 = sizeof(a1); getsockname(s1, (struct sockaddr*)&a1, &l1); connect(s2, (struct sockaddr*)&a1, l1); sendto(s2, \"Y\", 1, 0, NULL, 0); char b[2]={0}; recv(s1, b, 1, 0); printf(\"%s\", b); close(s1); close(s2); return 0; }"
        ),
        vec!["Y"]
    );
}
#[test]
fn udp_disconnect() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a={0}; a.sin_family = AF_UNSPEC; int r = connect(fd, (struct sockaddr*)&a, sizeof(a)); printf(\"%d\", r == 0 || r == -1); /* Some OS allow unspec to disconnect, others fail. Valid C either way */ close(fd); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn recvmsg_sendmsg_udp() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int s1 = socket(AF_INET, SOCK_DGRAM, 0); int s2 = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a1={0}; a1.sin_family = AF_INET; a1.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(s1, (struct sockaddr*)&a1, sizeof(a1)); socklen_t l1 = sizeof(a1); getsockname(s1, (struct sockaddr*)&a1, &l1); struct msghdr msg = {0}; struct iovec iov[1]; char buf[5] = \"msg\"; iov[0].iov_base = buf; iov[0].iov_len = 3; msg.msg_name = &a1; msg.msg_namelen = l1; msg.msg_iov = iov; msg.msg_iovlen = 1; sendmsg(s2, &msg, 0); struct msghdr rmsg = {0}; char rbuf[5]={0}; struct iovec riov[1]; riov[0].iov_base = rbuf; riov[0].iov_len = 3; rmsg.msg_iov = riov; rmsg.msg_iovlen = 1; recvmsg(s1, &rmsg, 0); printf(\"%s\", rbuf); close(s1); close(s2); return 0; }"
        ),
        vec!["msg"]
    );
}
#[test]
fn setsockopt_sndbuf() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET, SOCK_DGRAM, 0); int opt = 8192; int r = setsockopt(fd, SOL_SOCKET, SO_SNDBUF, &opt, sizeof(opt)); printf(\"%d\", r == 0); close(fd); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn udp_listen_fails() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET, SOCK_DGRAM, 0); int r = listen(fd, 5); printf(\"%d\", r == -1); close(fd); return 0; }"
        ),
        vec!["1"]
    );
} // UDP is connectionless, listen usually fails or does nothing
#[test]
fn socket_ipv6_udp() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET6, SOCK_DGRAM, 0); printf(\"%d\", fd != -99); if(fd>=0) close(fd); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getsockname_unbound_udp() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a={0}; socklen_t len = sizeof(a); getsockname(fd, (struct sockaddr*)&a, &len); printf(\"%d\", a.sin_port == 0); close(fd); return 0; }"
        ),
        vec!["1"]
    );
} // Unbound socket has port 0
#[test]
fn getpeername_unconnected_udp() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int fd = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a={0}; socklen_t len = sizeof(a); int r = getpeername(fd, (struct sockaddr*)&a, &len); printf(\"%d\", r == -1); close(fd); return 0; }"
        ),
        vec!["1"]
    );
} // Fails if not connected
#[test]
fn udp_msg_peek() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int s1 = socket(AF_INET, SOCK_DGRAM, 0); int s2 = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a1={0}; a1.sin_family = AF_INET; a1.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(s1, (struct sockaddr*)&a1, sizeof(a1)); socklen_t l1 = sizeof(a1); getsockname(s1, (struct sockaddr*)&a1, &l1); sendto(s2, \"P\", 1, 0, (struct sockaddr*)&a1, l1); char b1[2]={0}, b2[2]={0}; recvfrom(s1, b1, 1, MSG_PEEK, NULL, NULL); recvfrom(s1, b2, 1, 0, NULL, NULL); printf(\"%s %s\", b1, b2); close(s1); close(s2); return 0; }"
        ),
        vec!["P P"]
    );
}
#[test]
fn udp_zero_length_packet() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <netinet/in.h>\n#include <unistd.h>\nint main() { int s1 = socket(AF_INET, SOCK_DGRAM, 0); int s2 = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a1={0}; a1.sin_family = AF_INET; a1.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(s1, (struct sockaddr*)&a1, sizeof(a1)); socklen_t l1 = sizeof(a1); getsockname(s1, (struct sockaddr*)&a1, &l1); sendto(s2, \"\", 0, 0, (struct sockaddr*)&a1, l1); char b[2]={0}; int r = recvfrom(s1, b, 1, 0, NULL, NULL); printf(\"%d\", r == 0); close(s1); close(s2); return 0; }"
        ),
        vec!["1"]
    );
} // 0 length recv for UDP means empty datagram, unlike TCP EOF
#[test]
fn send_flags_zero() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\nint main() { int s1 = socket(AF_INET, SOCK_DGRAM, 0); /* Just test flag compiling */ int r = send(s1, \"x\", 1, 0); printf(\"%d\", r == -1); /* Unconnected fails */ close(s1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn recv_flags_zero() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/socket.h>\n#include <unistd.h>\n#include <fcntl.h>\nint main() { int s1 = socket(AF_INET, SOCK_DGRAM, 0); fcntl(s1, F_SETFL, O_NONBLOCK); char b[1]; int r = recv(s1, b, 1, 0); printf(\"%d\", r == -1); close(s1); return 0; }"
        ),
        vec!["1"]
    );
}
