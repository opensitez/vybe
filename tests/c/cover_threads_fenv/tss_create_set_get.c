// vybe-test: c/cover_threads_fenv/tss_create_set_get
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <threads.h>
int main() {
tss_t key; tss_create(&key, 0); tss_set(key, (void*)1); tss_get(key); tss_delete(key); return 0;
}

