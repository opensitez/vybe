// vybe-test: go/cover_os_syscall/syscall_mmap
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
func main() { _, _, _ = syscall.Mmap(-1, 0, 4096, syscall.PROT_READ, syscall.MAP_ANON) }
