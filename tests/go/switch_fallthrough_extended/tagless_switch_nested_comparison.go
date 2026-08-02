// vybe-test: go/switch_fallthrough_extended/tagless_switch_nested_comparison
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { score := 85
switch { case score >= 90: fmt.Println("A")
case score >= 80: fmt.Println("B")
default: fmt.Println("C") } }
