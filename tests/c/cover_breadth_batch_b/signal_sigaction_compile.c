// vybe-test: c/cover_breadth_batch_b/signal_sigaction_compile
// origin: languages/c/tests/c/test_cover_breadth_batch_b.rs
// vybe-test-mode: compile
#include <signal.h>
int main() {
struct sigaction sa={0}; sigaction(SIGINT,&sa,0); return 0;
}

