// vybe-test: go/switch_fallthrough_extended/switch_on_float_not_equal_use_tagless
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { f := 1.5
switch { case f == 1.5: fmt.Println("eq")
default: fmt.Println("ne") } }
