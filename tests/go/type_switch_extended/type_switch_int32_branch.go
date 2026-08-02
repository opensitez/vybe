// vybe-test: go/type_switch_extended/type_switch_int32_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case int32: fmt.Println("i32") default: fmt.Println("other") } }
func main() { tag(int32(-2)) }
