// vybe-test: go/switch_fallthrough_extended/switch_on_byte
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch byte(10) { case 10: fmt.Println("ten")
default: fmt.Println("other") } }
