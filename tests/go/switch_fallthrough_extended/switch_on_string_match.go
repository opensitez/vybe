// vybe-test: go/switch_fallthrough_extended/switch_on_string_match
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch "go" { case "go": fmt.Println("match")
default: fmt.Println("miss") } }
