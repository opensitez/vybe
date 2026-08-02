// vybe-test: go/type_switch_extended/type_switch_interface_subcase_via_type_assert
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type fmtStringer interface { String() string }
type myInt int
func (m myInt) String() string { return "m" }
func tag(v interface{}) { switch v.(type) { case fmtStringer: fmt.Println(v.(fmtStringer).String())
default: fmt.Println("no") } }
func main() { tag(myInt(1)) }
