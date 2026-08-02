// vybe-test: go/panic_recover_rules/recover_in_loop_defer_each_iteration
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { for i := 0; i < 2; i++ { defer func(n int) { if recover() != nil { fmt.Println(n) } }(i) }
panic("loop") }
func main() { defer func() { recover() }()
run() }
