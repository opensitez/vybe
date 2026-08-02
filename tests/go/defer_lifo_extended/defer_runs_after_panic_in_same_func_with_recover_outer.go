// vybe-test: go/defer_lifo_extended/defer_runs_after_panic_in_same_func_with_recover_outer
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer fmt.Println("a")
panic("p") }
func main() { defer func() { recover() }()
run() }
