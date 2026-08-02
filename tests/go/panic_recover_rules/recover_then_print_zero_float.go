// vybe-test: go/panic_recover_rules/recover_then_print_zero_float
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { recover()
var f float64
fmt.Println(f) }()
panic(true) }
func main() { run() }
