// vybe-test: go/switch_type_tagless/type_switch_string_branch
// origin: languages/go/tests/go/test_switch_type_tagless.rs

package main
import "fmt"
func describe(v interface{}) { switch v.(type) { case string: fmt.Println("str") default: fmt.Println("other") } }
func main() { describe("x") }
