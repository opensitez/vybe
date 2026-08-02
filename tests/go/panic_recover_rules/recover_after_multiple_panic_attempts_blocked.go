// vybe-test: go/panic_recover_rules/recover_after_multiple_panic_attempts_blocked
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover())
fmt.Println(recover() == nil) }()
panic("first") }
func main() { run() }
