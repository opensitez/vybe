// vybe-test: go/panic_recover_rules/defer_recover_prints_after_deferred_work
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer fmt.Println("late")
defer func() { if recover() != nil { fmt.Println("saved") } }()
panic("fail") }
func main() { run() }
