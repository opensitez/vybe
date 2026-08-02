// vybe-test: go/type_switch_extended/type_switch_concrete_string_regular_switch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func main() { s := "ab"
switch s { case "ab": fmt.Println("match")
default: fmt.Println("miss") } }
