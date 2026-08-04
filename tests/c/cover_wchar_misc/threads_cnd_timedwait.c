// vybe-test: c/cover_wchar_misc/threads_cnd_timedwait
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <threads.h>
int main() {
cnd_t c; mtx_t m; cnd_init(&c); mtx_init(&m,mtx_plain); cnd_timedwait(&c,&m,0); cnd_destroy(&c); mtx_destroy(&m); return 0;
}

