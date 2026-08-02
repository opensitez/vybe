// vybe-test: go/cover_os_syscall/signal_ignored
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "os/signal"
import "syscall"
func main() { _ = signal.Ignored(syscall.SIGPIPE) }
