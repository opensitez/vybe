// vybe-test: go/panic_recover_rules/nested_defer_outer_recover_gets_panic
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
defer func() { panic("inner") }() }
func main() { run() }
