// vybe-test: go/cover_os_syscall/user_name_field
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "os/user"
func main() { u, _ := user.Current()
if u != nil { _ = u.Name } }
