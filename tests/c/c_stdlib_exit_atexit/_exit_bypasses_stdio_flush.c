// vybe-test: c/c_stdlib_exit_atexit/_exit_bypasses_stdio_flush
// origin: languages/c/tests/c/test_c_stdlib_exit_atexit.rs
#include <stdio.h>
#include <stdlib.h>
#include <unistd.h>
int main() {
    printf("hello");
    _exit(0);
    return 0;
}

