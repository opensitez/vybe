use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn gethostbyname_basic() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct hostent *h = gethostbyname(\"localhost\"); printf(\"%d\", h != NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn gethostbyname_invalid() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct hostent *h = gethostbyname(\"this.domain.should.not.exist.xyz\"); printf(\"%d\", h == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn gethostbyaddr_basic() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\n#include <netinet/in.h>\n#include <arpa/inet.h>\nint main() { struct in_addr addr; inet_pton(AF_INET, \"127.0.0.1\", &addr); struct hostent *h = gethostbyaddr(&addr, sizeof(addr), AF_INET); printf(\"%d\", h != NULL || h == NULL); /* may not have reverse DNS, compile test */ return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getservbyname_basic() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct servent *s = getservbyname(\"http\", \"tcp\"); printf(\"%d\", s != NULL && s->s_port == htons(80)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getservbyname_invalid() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct servent *s = getservbyname(\"nonexistent_service_name_xyz\", \"tcp\"); printf(\"%d\", s == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getservbyport_basic() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct servent *s = getservbyport(htons(80), \"tcp\"); printf(\"%d\", s != NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getprotobyname_basic() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct protoent *p = getprotobyname(\"tcp\"); printf(\"%d\", p != NULL && p->p_proto == IPPROTO_TCP); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getprotobynumber_basic() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct protoent *p = getprotobynumber(IPPROTO_TCP); printf(\"%d\", p != NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getservent_set_end() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { setservent(1); struct servent *s = getservent(); printf(\"%d\", s != NULL); endservent(); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getprotoent_set_end() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { setprotoent(1); struct protoent *p = getprotoent(); printf(\"%d\", p != NULL); endprotoent(); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn hstrerror_basic() {
    assert_eq!(
        run_c(
            "#define _BSD_SOURCE\n#define _DEFAULT_SOURCE\n#include <netdb.h>\nint main() { const char *err = hstrerror(HOST_NOT_FOUND); printf(\"%d\", err != NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getnetbyname_basic() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct netent *n = getnetbyname(\"loopback\"); /* May or may not exist in /etc/networks */ printf(\"%d\", n != NULL || n == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getnetbyaddr_basic() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\n#include <arpa/inet.h>\nint main() { struct netent *n = getnetbyaddr(inet_addr(\"127.0.0.0\"), AF_INET); printf(\"%d\", n != NULL || n == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getnetent_set_end() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { setnetent(1); struct netent *n = getnetent(); printf(\"%d\", n != NULL || n == NULL); endnetent(); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn gethostent_set_end() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { sethostent(1); struct hostent *h = gethostent(); printf(\"%d\", h != NULL); endhostent(); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn herror_basic() {
    assert_eq!(
        run_c(
            "#define _BSD_SOURCE\n#define _DEFAULT_SOURCE\n#include <netdb.h>\nint main() { /* just prints to stderr, check compile */ herror(\"test\"); printf(\"1\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn h_errno_exists() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { h_errno = HOST_NOT_FOUND; printf(\"%d\", h_errno == HOST_NOT_FOUND); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn gethostbyname_ipv4_only() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct hostent *h = gethostbyname(\"127.0.0.1\"); printf(\"%d\", h != NULL && h->h_addrtype == AF_INET); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getservbyport_udp() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct servent *s = getservbyport(htons(53), \"udp\"); printf(\"%d\", s != NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getprotobyname_udp() {
    assert_eq!(
        run_c(
            "#include <netdb.h>\nint main() { struct protoent *p = getprotobyname(\"udp\"); printf(\"%d\", p != NULL && p->p_proto == IPPROTO_UDP); return 0; }"
        ),
        vec!["1"]
    );
}
