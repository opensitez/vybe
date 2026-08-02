// vybe-test: go/defer_panic_recover_extra/recover_returns_value_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
panic("boom") }
func main() { run() }
