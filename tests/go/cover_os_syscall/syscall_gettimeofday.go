// vybe-test: go/cover_os_syscall/syscall_gettimeofday
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
type Timeval = syscall.Timeval
func main() { var tv Timeval
_ = syscall.Gettimeofday(&tv) }
