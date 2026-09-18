// vybe-test: c/cover_breadth_batch_a/signal_kill_compile
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <signal.h>
#include <unistd.h>
int main() {
return kill(getpid(), 0);
}

