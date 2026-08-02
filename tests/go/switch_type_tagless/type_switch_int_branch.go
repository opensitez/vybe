// vybe-test: go/switch_type_tagless/type_switch_int_branch
// origin: languages/go/tests/go/test_switch_type_tagless.rs

package main
import "fmt"
func describe(v interface{}) { switch v.(type) { case int: fmt.Println("int") default: fmt.Println("other") } }
func main() { describe(3) }
