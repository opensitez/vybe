// vybe-test: go/panic_recover_rules/recover_not_effective_in_same_func_as_panic
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() {}()
panic("x")
_ = recover() }
func main() { defer func() { recover() }()
defer func() { fmt.Println("shield") }()
defer func() { recover() }()
func() { defer func() { if recover() != nil { fmt.Println("caught") } }()
panic("x") }() }
