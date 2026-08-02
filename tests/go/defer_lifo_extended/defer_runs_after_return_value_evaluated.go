// vybe-test: go/defer_lifo_extended/defer_runs_after_return_value_evaluated
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() int { defer fmt.Println("defer")
return func() int { fmt.Println("ret")
return 2 }() }
func main() { fmt.Println(work()) }
