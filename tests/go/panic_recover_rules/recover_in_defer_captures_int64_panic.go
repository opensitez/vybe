// vybe-test: go/panic_recover_rules/recover_in_defer_captures_int64_panic
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
panic(int64(99)) }
func main() { run() }
