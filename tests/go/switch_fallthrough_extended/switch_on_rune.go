// vybe-test: go/switch_fallthrough_extended/switch_on_rune
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch rune('A') { case 'A': fmt.Println("A")
default: fmt.Println("other") } }
