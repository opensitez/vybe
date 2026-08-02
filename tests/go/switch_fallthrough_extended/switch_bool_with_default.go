// vybe-test: go/switch_fallthrough_extended/switch_bool_with_default
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch bool(0 == 1) { default: fmt.Println("def")
case true: fmt.Println("t") } }
