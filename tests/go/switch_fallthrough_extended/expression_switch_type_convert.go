// vybe-test: go/switch_fallthrough_extended/expression_switch_type_convert
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
type myint int
func main() { switch myint(3) { case 3: fmt.Println("n")
default: fmt.Println("o") } }
