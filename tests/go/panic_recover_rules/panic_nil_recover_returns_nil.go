// vybe-test: go/panic_recover_rules/panic_nil_recover_returns_nil
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover() == nil) }()
panic(nil) }
func main() { run() }
