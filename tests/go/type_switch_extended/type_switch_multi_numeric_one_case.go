// vybe-test: go/type_switch_extended/type_switch_multi_numeric_one_case
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case int, int32, int64: fmt.Println("num") default: fmt.Println("other") } }
func main() { tag(int32(3)) }
