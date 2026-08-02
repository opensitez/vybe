// vybe-test: go/cover_os_syscall/syscall_wait4
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
type WaitStatus = syscall.WaitStatus
func main() { var status WaitStatus
_, _ = syscall.Wait4(-1, &status, syscall.WNOHANG, nil) }
