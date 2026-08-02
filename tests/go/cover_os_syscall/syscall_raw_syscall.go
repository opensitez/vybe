// vybe-test: go/cover_os_syscall/syscall_raw_syscall
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
func main() { _, _, _ = syscall.RawSyscall(syscall.SYS_GETPID, 0, 0, 0) }
