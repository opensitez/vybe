// vybe-test: go/panic_recover_rules/recover_then_print_zero_int
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { recover()
var n int
fmt.Println(n) }()
panic("stop") }
func main() { run() }
