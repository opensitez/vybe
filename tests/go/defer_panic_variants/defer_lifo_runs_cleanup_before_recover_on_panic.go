// vybe-test: go/defer_panic_variants/defer_lifo_runs_cleanup_before_recover_on_panic
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer fmt.Println("cleanup")
defer func() { recover() }()
panic("stop") }
func main() { run()
fmt.Println("done") }
