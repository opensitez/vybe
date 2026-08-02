// vybe-test: go/type_switch_extended/type_switch_multi_float_case
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case float32, float64: fmt.Println("float")
default: fmt.Println("other") } }
func main() { tag(float64(2.0)) }
