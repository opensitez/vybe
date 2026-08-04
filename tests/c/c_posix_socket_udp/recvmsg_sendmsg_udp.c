// vybe-test: c/c_posix_socket_udp/recvmsg_sendmsg_udp
// origin: languages/c/tests/c/test_c_posix_socket_udp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <netinet/in.h>
#include <unistd.h>
int main() {const char *__w[] = {"msg"};
int __n = 1, __i = 0;
 int s1 = socket(AF_INET, SOCK_DGRAM, 0); int s2 = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a1={0}; a1.sin_family = AF_INET; a1.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(s1, (struct sockaddr*)&a1, sizeof(a1)); socklen_t l1 = sizeof(a1); getsockname(s1, (struct sockaddr*)&a1, &l1); struct msghdr msg = {0}; struct iovec iov[1]; char buf[5] = "msg"; iov[0].iov_base = buf; iov[0].iov_len = 3; msg.msg_name = &a1; msg.msg_namelen = l1; msg.msg_iov = iov; msg.msg_iovlen = 1; sendmsg(s2, &msg, 0); struct msghdr rmsg = {0}; char rbuf[5]={0}; struct iovec riov[1]; riov[0].iov_base = rbuf; riov[0].iov_len = 3; rmsg.msg_iov = riov; rmsg.msg_iovlen = 1; recvmsg(s1, &rmsg, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%s", rbuf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(s1); close(s2); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

