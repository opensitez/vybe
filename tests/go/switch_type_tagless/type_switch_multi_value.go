// vybe-test: go/switch_type_tagless/type_switch_multi_value
// origin: languages/go/tests/go/test_switch_type_tagless.rs

package main
import "fmt"
func describe(v interface{}) { switch v.(type) { case int, int64: fmt.Println("integer") default: fmt.Println("other") } }
func main() { describe(int64(9)) }
