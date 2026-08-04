// vybe-test: c/cover_wchar_misc/threads_mtx_timedlock
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <threads.h>
int main() {
mtx_t m; mtx_init(&m,mtx_plain); mtx_timedlock(&m,0); mtx_unlock(&m); mtx_destroy(&m); return 0;
}

