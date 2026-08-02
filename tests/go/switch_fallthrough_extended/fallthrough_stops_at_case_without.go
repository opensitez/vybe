// vybe-test: go/switch_fallthrough_extended/fallthrough_stops_at_case_without
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch 1 { case 1: fmt.Println(1)
fallthrough
case 2: fmt.Println(2)
case 3: fmt.Println(3) } }
