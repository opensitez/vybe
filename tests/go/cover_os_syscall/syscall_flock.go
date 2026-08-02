// vybe-test: go/cover_os_syscall/syscall_flock
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
func main() { fd, _ := syscall.Open(".", syscall.O_RDONLY, 0)
if fd >= 0 { defer syscall.Close(fd)
_ = syscall.Flock(fd, syscall.LOCK_EX) } }
