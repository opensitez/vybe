// vybe-test: go/type_switch_extended/type_switch_uint_in_multi_int_case
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case int, uint, int64: fmt.Println("integer")
default: fmt.Println("other") } }
func main() { tag(uint(9)) }
