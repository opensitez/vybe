// vybe-test: go/switch_fallthrough_extended/switch_on_int_negative
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch -3 { case -3: fmt.Println("neg")
default: fmt.Println("other") } }
