// vybe-test: c/cover_breadth_batch_b/signal_pthread_sigmask
// origin: languages/c/tests/c/test_cover_breadth_batch_b.rs
// vybe-test-mode: compile
#include <signal.h>
int main() {
sigset_t s; sigemptyset(&s); return 0;
}

