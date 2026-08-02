// vybe-test: go/switch_fallthrough_extended/tagless_switch_multiple_false_then_true
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { n := 12
switch { case n%2 == 1: fmt.Println("odd")
case n%3 == 0: fmt.Println("three")
case n%4 == 0: fmt.Println("four")
default: fmt.Println("none") } }
