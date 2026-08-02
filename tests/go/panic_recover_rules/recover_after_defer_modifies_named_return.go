// vybe-test: go/panic_recover_rules/recover_after_defer_modifies_named_return
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() (n int) { defer func() { recover()
n = 9 }()
panic("p")
return 1 }
func main() { fmt.Println(run()) }
