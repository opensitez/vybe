// vybe-test: go/switch_fallthrough_extended/switch_on_string_default
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch "rust" { case "go": fmt.Println("go")
default: fmt.Println("other") } }
