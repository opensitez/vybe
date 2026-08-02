// vybe-test: go/defer_lifo_extended/defer_runs_at_function_exit_not_block
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { if true { defer fmt.Println("defer")
fmt.Println("block")
}
fmt.Println("main") }
