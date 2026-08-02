// vybe-test: go/panic_recover_rules/recover_nil_when_no_panic_in_defer
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover() == nil) }() }
func main() { run() }
