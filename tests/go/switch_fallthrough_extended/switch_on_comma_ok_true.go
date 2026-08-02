// vybe-test: go/switch_fallthrough_extended/switch_on_comma_ok_true
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { m := map[string]int{"a": 1}
switch v, ok := m["a"]; ok { case true: fmt.Println(v)
default: fmt.Println(0) } }
