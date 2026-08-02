// vybe-test: go/defer_lifo_extended/defer_after_panic_still_runs_with_recover
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer fmt.Println("cleanup")
defer func() { recover() }()
panic("stop")
fmt.Println("skip") }
func main() { run()
fmt.Println("done") }
