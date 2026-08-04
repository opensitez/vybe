// vybe-test: c/cover_wchar_misc/threads_thrd_detach
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <threads.h>
int w(void *a){(void)a;return 0;}
int main() {
thrd_t t; thrd_create(&t,w,0); thrd_detach(t); return 0;
}

