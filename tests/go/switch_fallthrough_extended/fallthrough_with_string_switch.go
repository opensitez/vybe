// vybe-test: go/switch_fallthrough_extended/fallthrough_with_string_switch
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch "x" { case "x": fmt.Println("a")
fallthrough
case "y": fmt.Println("b")
default: fmt.Println("c") } }
