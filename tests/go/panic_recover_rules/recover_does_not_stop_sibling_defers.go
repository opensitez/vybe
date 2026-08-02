// vybe-test: go/panic_recover_rules/recover_does_not_stop_sibling_defers
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer fmt.Println("sibling")
defer func() { recover() }()
panic("p") }
func main() { run() }
