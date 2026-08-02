// vybe-test: go/type_switch_extended/type_switch_multi_string_rune_case
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case string, rune: fmt.Println("text") default: fmt.Println("other") } }
func main() { tag('Z') }
