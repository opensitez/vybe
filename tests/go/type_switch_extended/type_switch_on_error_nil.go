// vybe-test: go/type_switch_extended/type_switch_on_error_nil
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case error: fmt.Println("err")
default: fmt.Println("nil-err") } }
func main() { var e error
tag(e) }
