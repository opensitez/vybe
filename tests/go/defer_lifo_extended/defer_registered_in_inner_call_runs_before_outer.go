// vybe-test: go/defer_lifo_extended/defer_registered_in_inner_call_runs_before_outer
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func inner() { defer fmt.Println("inner")
}
func main() { defer fmt.Println("outer")
inner()
}
