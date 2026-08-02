// vybe-test: go/switch_fallthrough_extended/switch_string_case_sensitive
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch "Go" { case "go": fmt.Println("lower")
case "Go": fmt.Println("exact")
default: fmt.Println("other") } }
