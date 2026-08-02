// vybe-test: go/switch_fallthrough_extended/tagless_switch_with_and_condition
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { a, b := 2, 3
switch { case a < 5 && b > 2: fmt.Println("ok")
default: fmt.Println("no") } }
