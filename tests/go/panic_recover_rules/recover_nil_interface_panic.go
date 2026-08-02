// vybe-test: go/panic_recover_rules/recover_nil_interface_panic
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover() == nil) }()
var p *int
panic(p) }
func main() { run() }
