// vybe-test: go/type_switch_extended/type_switch_func_with_params
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case func(int) int: fmt.Println("fn")
default: fmt.Println("other") } }
func main() { tag(func(x int) int { return x }) }
