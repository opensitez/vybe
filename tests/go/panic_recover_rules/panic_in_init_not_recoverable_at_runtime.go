// vybe-test: go/panic_recover_rules/panic_in_init_not_recoverable_at_runtime
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover() == nil) }() }
func main() { run() }
