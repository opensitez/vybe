// vybe-test: go/type_switch_extended/type_switch_float32_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case float32: fmt.Println("f32") default: fmt.Println("other") } }
func main() { tag(float32(1.25)) }
