// vybe-test: go/type_switch_extended/type_switch_error_interface_value
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
import "errors"
func tag(v interface{}) { switch v.(type) { case error: fmt.Println(v.Error())
default: fmt.Println("no") } }
func main() { tag(errors.New("fail")) }
