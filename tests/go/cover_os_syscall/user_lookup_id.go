// vybe-test: go/cover_os_syscall/user_lookup_id
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "os/user"
func main() { _, _ = user.LookupId("0") }
