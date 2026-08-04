// vybe-test: c/c_posix_message_queues/mq_setattr
// origin: languages/c/tests/c/test_c_posix_message_queues.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <mqueue.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 mqd_t q = mq_open("/test_mq5", O_CREAT | O_RDWR, 0644, NULL); if(q != (mqd_t)-1) { struct mq_attr a = {0}; a.mq_flags = O_NONBLOCK; struct mq_attr old; int r = mq_setattr(q, &a, &old); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } mq_close(q); mq_unlink("/test_mq5"); } else { char __t[512]; snprintf(__t, sizeof(__t), "1");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

