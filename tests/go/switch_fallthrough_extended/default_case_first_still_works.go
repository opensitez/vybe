// vybe-test: go/switch_fallthrough_extended/default_case_first_still_works
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch 5 { default: fmt.Println("def")
case 1: fmt.Println(1) } }
