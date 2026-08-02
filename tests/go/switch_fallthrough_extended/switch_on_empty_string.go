// vybe-test: go/switch_fallthrough_extended/switch_on_empty_string
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch "" { case "": fmt.Println("empty")
default: fmt.Println("other") } }
