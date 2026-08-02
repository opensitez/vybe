// vybe-test: go/type_switch_extended/type_switch_concrete_int_regular_switch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func main() { x := 2
switch x { case 1: fmt.Println("one")
case 2: fmt.Println("two")
default: fmt.Println("other") } }
