// vybe-test: go/panic_recover_rules/recover_consumes_panic_once
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover())
fmt.Println(recover() == nil) }()
panic("once") }
func main() { run() }
