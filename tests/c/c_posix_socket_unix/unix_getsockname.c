// vybe-test: c/c_posix_socket_unix/unix_getsockname
// origin: languages/c/tests/c/test_c_posix_socket_unix.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <sys/un.h>
#include <unistd.h>
#include <string.h>
int main() {const char *__w[] = {"test_unix4.sock"};
int __n = 1, __i = 0;
 unlink("test_unix4.sock"); int fd = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; strcpy(addr.sun_path, "test_unix4.sock"); bind(fd, (struct sockaddr*)&addr, sizeof(addr)); struct sockaddr_un a2={0}; socklen_t len = sizeof(a2); getsockname(fd, (struct sockaddr*)&a2, &len); { char __t[512]; snprintf(__t, sizeof(__t), "%s", a2.sun_path);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd); unlink("test_unix4.sock"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

