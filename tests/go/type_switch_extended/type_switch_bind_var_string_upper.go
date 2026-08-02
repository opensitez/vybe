// vybe-test: go/type_switch_extended/type_switch_bind_var_string_upper
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func work(v interface{}) { switch s := v.(type) { case string: fmt.Println(s + "!") default: fmt.Println("skip") } }
func main() { work("go") }
