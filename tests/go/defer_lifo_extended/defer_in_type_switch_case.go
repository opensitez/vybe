// vybe-test: go/defer_lifo_extended/defer_in_type_switch_case
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case int: defer fmt.Println("int")
default: defer fmt.Println("def") } }
func main() { tag(1) }
