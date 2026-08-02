// vybe-test: go/type_switch_extended/type_switch_bind_var_slice_cap
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func work(v interface{}) { switch s := v.(type) { case []int: fmt.Println(len(s))
fmt.Println(cap(s)) default: fmt.Println(0) } }
func main() { work([]int{1, 2, 3}) }
