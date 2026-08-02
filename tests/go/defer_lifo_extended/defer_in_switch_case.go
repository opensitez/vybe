// vybe-test: go/defer_lifo_extended/defer_in_switch_case
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { switch 2 { case 2: defer fmt.Println("case")
default: defer fmt.Println("def") } }
