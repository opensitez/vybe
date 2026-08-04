// vybe-test: c/c_posix_poll_select/select_all_null
// origin: languages/c/tests/c/test_c_posix_poll_select.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/select.h>
int main() { int res = select(0, NULL, NULL, NULL, NULL); /* blocks forever normally, but we test signal or just compile. Wait, we don't want to block forever. Use timeout */ struct timeval tv = {0, 0}; res = select(0, NULL, NULL, NULL, &tv); printf("%d", res == 0); return 0; }

