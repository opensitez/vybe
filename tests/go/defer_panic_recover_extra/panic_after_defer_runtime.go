// vybe-test: go/defer_panic_recover_extra/panic_after_defer_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func run() { defer fmt.Println("cleanup")
defer func() { recover() }()
panic("stop") }
func main() { run() }
