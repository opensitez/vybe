// vybe-test: go/defer_lifo_extended/defer_before_return_zero_value
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() int { defer fmt.Println("d")
return 0 }
func main() { fmt.Println(work()) }
