// vybe-test: go/type_switch_extended/type_switch_defer_wraps_switch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { defer fmt.Println("end")
switch v.(type) { case int: fmt.Println("int")
default: fmt.Println("def") } }
func main() { tag(1) }
