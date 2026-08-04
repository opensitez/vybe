// vybe-test: c/c_posix_socket_unix/unix_dgram_sendto
// origin: languages/c/tests/c/test_c_posix_socket_unix.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <sys/un.h>
#include <unistd.h>
#include <string.h>
int main() {const char *__w[] = {"D"};
int __n = 1, __i = 0;
 unlink("t1.sock"); unlink("t2.sock"); int s1 = socket(AF_UNIX, SOCK_DGRAM, 0); int s2 = socket(AF_UNIX, SOCK_DGRAM, 0); struct sockaddr_un a1={0}, a2={0}; a1.sun_family = AF_UNIX; strcpy(a1.sun_path, "t1.sock"); a2.sun_family = AF_UNIX; strcpy(a2.sun_path, "t2.sock"); bind(s1, (struct sockaddr*)&a1, sizeof(a1)); bind(s2, (struct sockaddr*)&a2, sizeof(a2)); sendto(s1, "D", 1, 0, (struct sockaddr*)&a2, sizeof(a2)); char b[2]={0}; recvfrom(s2, b, 1, 0, NULL, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%s", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(s1); close(s2); unlink("t1.sock"); unlink("t2.sock"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

