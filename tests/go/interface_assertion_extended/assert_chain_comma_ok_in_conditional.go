// vybe-test: go/interface_assertion_extended/assert_chain_comma_ok_in_conditional
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func pick(v interface{}) { if s, ok := v.(string); ok { fmt.Println(s) } else if n, ok := v.(int); ok { fmt.Println(n) } else { fmt.Println("none") } }
func main() { pick(3) }
