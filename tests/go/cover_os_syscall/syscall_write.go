// vybe-test: go/cover_os_syscall/syscall_write
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
func main() { _, _ = syscall.Write(syscall.Stdout, []byte("x")) }
