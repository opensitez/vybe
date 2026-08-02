// vybe-test: go/type_switch_extended/type_switch_byte_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case byte: fmt.Println("byte") default: fmt.Println("other") } }
func main() { tag(byte(65)) }
