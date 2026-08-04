// vybe-test: c/c_posix_message_queues/mq_open_exclusive
// origin: languages/c/tests/c/test_c_posix_message_queues.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <mqueue.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 mqd_t q1 = mq_open("/test_mq6", O_CREAT | O_RDWR, 0644, NULL); if(q1 != (mqd_t)-1) { mqd_t q2 = mq_open("/test_mq6", O_CREAT | O_EXCL | O_RDWR, 0644, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", q2 == (mqd_t)-1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } mq_close(q1); mq_unlink("/test_mq6"); } else { char __t[512]; snprintf(__t, sizeof(__t), "1");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

