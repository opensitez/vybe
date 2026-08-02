// vybe-test: go/cover_os_syscall/syscall_bind
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
func main() { fd, _ := syscall.Socket(syscall.AF_INET, syscall.SOCK_STREAM, 0)
if fd >= 0 { defer syscall.Close(fd)
sa := &syscall.SockaddrInet4{Port: 0}
_ = syscall.Bind(fd, sa) } }
