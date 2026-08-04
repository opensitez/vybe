// vybe-test: c/c_posix_socket_udp/sendto_recvfrom_udp
// origin: languages/c/tests/c/test_c_posix_socket_udp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <netinet/in.h>
#include <unistd.h>
#include <pthread.h>
int port = 0;
void* f(void* a) { int s = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in addr={0}; addr.sin_family = AF_INET; addr.sin_addr.s_addr = htonl(INADDR_LOOPBACK); addr.sin_port = htons(port); sendto(s, "udp", 3, 0, (struct sockaddr*)&addr, sizeof(addr)); close(s); return NULL; }
int main() {const char *__w[] = {"udp"};
int __n = 1, __i = 0;
 int l = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a={0}; a.sin_family = AF_INET; a.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(l, (struct sockaddr*)&a, sizeof(a)); socklen_t len=sizeof(a); getsockname(l, (struct sockaddr*)&a, &len); port = ntohs(a.sin_port); pthread_t t; pthread_create(&t, NULL, f, NULL); char b[5]={0}; struct sockaddr_in from; socklen_t flen = sizeof(from); recvfrom(l, b, 3, 0, (struct sockaddr*)&from, &flen); { char __t[512]; snprintf(__t, sizeof(__t), "%s", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(l); pthread_join(t, NULL); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

