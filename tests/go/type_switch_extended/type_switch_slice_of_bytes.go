// vybe-test: go/type_switch_extended/type_switch_slice_of_bytes
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case []byte: fmt.Println(len(v.([]byte)))
default: fmt.Println(0) } }
func main() { tag([]byte{10, 20}) }
