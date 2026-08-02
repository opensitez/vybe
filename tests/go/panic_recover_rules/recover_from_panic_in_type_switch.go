// vybe-test: go/panic_recover_rules/recover_from_panic_in_type_switch
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
switch 0 { default: panic("sw") } }
func main() { run() }
