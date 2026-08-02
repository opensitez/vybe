// vybe-test: go/type_switch_extended/type_switch_float64_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case float64: fmt.Println("f64") default: fmt.Println("other") } }
func main() { tag(3.5) }
