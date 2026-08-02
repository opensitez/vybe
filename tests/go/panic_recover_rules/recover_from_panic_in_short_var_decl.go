// vybe-test: go/panic_recover_rules/recover_from_panic_in_short_var_decl
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
a := 1
_ = a
panic("decl") }
func main() { run() }
