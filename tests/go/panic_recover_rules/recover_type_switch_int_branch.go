// vybe-test: go/panic_recover_rules/recover_type_switch_int_branch
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { switch recover().(type) { case int: fmt.Println("int")
default: fmt.Println("other") } }()
panic(3) }
func main() { run() }
