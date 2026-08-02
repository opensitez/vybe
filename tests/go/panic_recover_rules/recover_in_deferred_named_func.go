// vybe-test: go/panic_recover_rules/recover_in_deferred_named_func
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func save() { fmt.Println(recover()) }
func run() { defer save()
panic("named") }
func main() { run() }
