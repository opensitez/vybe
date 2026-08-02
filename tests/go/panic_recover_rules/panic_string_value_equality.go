// vybe-test: go/panic_recover_rules/panic_string_value_equality
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { v := recover()
fmt.Println(v == "msg") }()
panic("msg") }
func main() { run() }
