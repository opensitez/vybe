// vybe-test: go/cover_os_syscall/user_current
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "os/user"
func main() { _, _ = user.Current() }
