// vybe-test: go/cover_os_syscall/signal_notify_context
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "context"
import "os/signal"
import "syscall"
func main() { ctx, stop := signal.NotifyContext(context.Background(), syscall.SIGINT)
defer stop()
_ = ctx }
