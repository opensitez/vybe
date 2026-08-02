// vybe-test: go/panic_recover_rules/recover_then_print_zero_string
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { recover()
var s string
fmt.Println(s == "") }()
panic(1) }
func main() { run() }
