// vybe-test: go/cover_os_syscall/syscall_errno
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
type Errno = syscall.Errno
func main() { var e Errno
_ = e.Error()
_ = syscall.ENOENT }
