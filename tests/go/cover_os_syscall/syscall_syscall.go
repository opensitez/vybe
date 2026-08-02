// vybe-test: go/cover_os_syscall/syscall_syscall
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
func main() { _, _, _ = syscall.Syscall(syscall.SYS_GETUID, 0, 0, 0) }
