// vybe-test: c/c_posix_daemon_sessions/daemon_func_gnu
// origin: languages/c/tests/c/test_c_posix_daemon_sessions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _BSD_SOURCE
#define _DEFAULT_SOURCE
#include <unistd.h>
int main() { /* don't actually daemonize the test runner */ printf("ok"); return 0; }

