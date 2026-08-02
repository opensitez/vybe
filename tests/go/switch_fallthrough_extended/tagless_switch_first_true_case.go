// vybe-test: go/switch_fallthrough_extended/tagless_switch_first_true_case
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { x := 7
switch { case x < 5: fmt.Println("low")
case x < 10: fmt.Println("mid")
default: fmt.Println("high") } }
