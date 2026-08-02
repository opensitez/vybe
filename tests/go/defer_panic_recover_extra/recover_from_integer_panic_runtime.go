// vybe-test: go/defer_panic_recover_extra/recover_from_integer_panic_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
panic(12) }
func main() { run() }
