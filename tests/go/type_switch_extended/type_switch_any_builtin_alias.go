// vybe-test: go/type_switch_extended/type_switch_any_builtin_alias
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v any) { switch v.(type) { case int: fmt.Println("any-int") default: fmt.Println("any-other") } }
func main() { tag(4) }
