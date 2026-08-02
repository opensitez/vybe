// vybe-test: go/switch_fallthrough_extended/tagless_switch_short_circuit
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { called := false
switch { case false: called = true
fmt.Println(1)
case true: fmt.Println(2) }
fmt.Println(called) }
