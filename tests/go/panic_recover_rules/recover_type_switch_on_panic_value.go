// vybe-test: go/panic_recover_rules/recover_type_switch_on_panic_value
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { switch recover().(type) { case string: fmt.Println("str")
case int: fmt.Println("int")
default: fmt.Println("other") } }()
panic("str") }
func main() { run() }
