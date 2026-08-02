// vybe-test: go/interface_assertion_extended/recover_after_assert_panic_allows_continue
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func main() { caught := false
func() { defer func() { if recover() != nil { caught = true } }()
var v interface{} = 1
_ = v.(bool) }()
if caught { fmt.Println("ok") } else { fmt.Println("no") } }
