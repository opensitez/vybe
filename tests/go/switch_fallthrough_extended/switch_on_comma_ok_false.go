// vybe-test: go/switch_fallthrough_extended/switch_on_comma_ok_false
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { m := map[string]int{}
switch _, ok := m["z"]; ok { case true: fmt.Println(1)
default: fmt.Println(0) } }
