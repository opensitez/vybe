// vybe-test: go/cover_os_syscall/syscall_byte_slice_to_string
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
func main() { _ = syscall.ByteSliceToString([]byte("go")) }
