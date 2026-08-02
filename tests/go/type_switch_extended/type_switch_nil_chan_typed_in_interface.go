// vybe-test: go/type_switch_extended/type_switch_nil_chan_typed_in_interface
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case chan string: fmt.Println("chan")
default: fmt.Println("nil-chan") } }
func main() { var c chan string
tag(c) }
