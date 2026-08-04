// vybe-test: c/c_posix_socket_unix/unix_abstract_namespace_gnu
// origin: languages/c/tests/c/test_c_posix_socket_unix.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <sys/socket.h>
#include <sys/un.h>
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; addr.sun_path[0] = '\0'; addr.sun_path[1] = 't'; addr.sun_path[2] = 's'; addr.sun_path[3] = 't'; int r = bind(fd, (struct sockaddr*)&addr, sizeof(sa_family_t) + 4); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0 || r != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

