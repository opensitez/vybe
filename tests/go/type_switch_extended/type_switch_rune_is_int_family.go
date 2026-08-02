// vybe-test: go/type_switch_extended/type_switch_rune_is_int_family
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case rune: fmt.Println("rune")
case int: fmt.Println("int")
default: fmt.Println("other") } }
func main() { tag(rune(65)) }
