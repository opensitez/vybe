// vybe-test: go/cover_os_syscall/user_group_name_field
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "os/user"
func main() { g, _ := user.LookupGroup("staff")
if g != nil { _ = g.Name } }
