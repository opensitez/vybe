// vybe-test: go/cover_os_syscall/signal_stop
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "os"
import "os/signal"
import "syscall"
func main() { ch := make(chan os.Signal, 1)
signal.Notify(ch, syscall.SIGTERM)
signal.Stop(ch) }
