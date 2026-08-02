// vybe-test: go/defer_lifo_extended/defer_before_return_in_nested_func
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func inner() int { defer fmt.Println("in")
return 1 }
func main() { defer fmt.Println("out")
fmt.Println(inner()) }
