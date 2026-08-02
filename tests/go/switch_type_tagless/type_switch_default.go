// vybe-test: go/switch_type_tagless/type_switch_default
// origin: languages/go/tests/go/test_switch_type_tagless.rs

package main
import "fmt"
func describe(v interface{}) { switch v.(type) { case int: fmt.Println("int") default: fmt.Println("default") } }
func main() { describe(1.5) }
