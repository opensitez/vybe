// vybe-test: go/type_switch_extended/type_switch_custom_interface_impl
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type speaker interface { Say() string }
type cat struct{}
func (c cat) Say() string { return "meow" }
func tag(v interface{}) { switch v.(type) { case speaker: fmt.Println(v.(speaker).Say())
default: fmt.Println("mute") } }
func main() { tag(cat{}) }
