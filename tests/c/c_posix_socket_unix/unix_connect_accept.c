// vybe-test: c/c_posix_socket_unix/unix_connect_accept
// origin: languages/c/tests/c/test_c_posix_socket_unix.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <sys/un.h>
#include <unistd.h>
#include <string.h>
#include <pthread.h>
void* f(void* a) { int s = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un addr={0}; addr.sun_family = AF_UNIX; strcpy(addr.sun_path, "test_unix3.sock"); while(connect(s, (struct sockaddr*)&addr, sizeof(addr)) != 0) usleep(10000); send(s, "unix", 4, 0); close(s); return NULL; }
int main() {const char *__w[] = {"unix"};
int __n = 1, __i = 0;
 unlink("test_unix3.sock"); int l = socket(AF_UNIX, SOCK_STREAM, 0); struct sockaddr_un a={0}; a.sun_family = AF_UNIX; strcpy(a.sun_path, "test_unix3.sock"); bind(l, (struct sockaddr*)&a, sizeof(a)); listen(l, 5); pthread_t t; pthread_create(&t, NULL, f, NULL); int c = accept(l, NULL, NULL); char b[5]={0}; recv(c, b, 4, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%s", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(c); close(l); pthread_join(t, NULL); unlink("test_unix3.sock"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

