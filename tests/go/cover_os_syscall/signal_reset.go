// vybe-test: go/cover_os_syscall/signal_reset
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "os/signal"
import "syscall"
func main() { signal.Reset(syscall.SIGHUP) }
