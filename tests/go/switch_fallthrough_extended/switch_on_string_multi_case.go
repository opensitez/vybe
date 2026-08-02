// vybe-test: go/switch_fallthrough_extended/switch_on_string_multi_case
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch "a" { case "a", "b": fmt.Println("ab")
default: fmt.Println("x") } }
