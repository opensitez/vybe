// vybe-test: go/switch_fallthrough_extended/tagless_switch_with_or_condition
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { x := 0
switch { case x == 1 || x == 0: fmt.Println("zero")
default: fmt.Println("other") } }
