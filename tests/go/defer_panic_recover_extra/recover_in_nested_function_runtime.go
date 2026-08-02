// vybe-test: go/defer_panic_recover_extra/recover_in_nested_function_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover() != nil) }()
panic("boom") }
func main() { run() }
