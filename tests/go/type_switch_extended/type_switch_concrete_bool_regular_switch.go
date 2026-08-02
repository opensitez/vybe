// vybe-test: go/type_switch_extended/type_switch_concrete_bool_regular_switch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func main() { switch true { case true: fmt.Println("yes")
case false: fmt.Println("no") } }
