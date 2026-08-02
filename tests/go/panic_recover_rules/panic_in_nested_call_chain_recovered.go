// vybe-test: go/panic_recover_rules/panic_in_nested_call_chain_recovered
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func leaf() { panic(7) }
func mid() { leaf() }
func run() { defer func() { fmt.Println(recover()) }()
mid() }
func main() { run() }
