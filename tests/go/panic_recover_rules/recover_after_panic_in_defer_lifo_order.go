// vybe-test: go/panic_recover_rules/recover_after_panic_in_defer_lifo_order
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println("a")
recover() }()
defer func() { fmt.Println("b")
panic("p") }() }
func main() { run() }
