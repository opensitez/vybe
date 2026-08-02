// vybe-test: go/panic_recover_rules/recover_then_print_zero_bool
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { recover()
var b bool
fmt.Println(b) }()
panic("z") }
func main() { run() }
