// vybe-test: go/cover_os_syscall/syscall_pipe
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
func main() { var p [2]int
_, _ = syscall.Pipe(p[:]) }
