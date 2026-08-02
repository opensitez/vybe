// vybe-test: go/switch_fallthrough_extended/tagless_switch_with_equality
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { a, b := 3, 3
switch { case a == b: fmt.Println("eq")
default: fmt.Println("ne") } }
