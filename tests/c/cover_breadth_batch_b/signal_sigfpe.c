// vybe-test: c/cover_breadth_batch_b/signal_sigfpe
// origin: languages/c/tests/c/test_cover_breadth_batch_b.rs
// vybe-test-mode: compile
#include <signal.h>
int main() {
return SIGFPE;
}

