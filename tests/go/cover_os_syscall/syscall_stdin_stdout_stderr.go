// vybe-test: go/cover_os_syscall/syscall_stdin_stdout_stderr
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
func main() { _ = syscall.Stdin
_ = syscall.Stdout
_ = syscall.Stderr }
