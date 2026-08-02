// vybe-test: go/panic_recover_rules/panic_in_anonymous_defer_recovered_by_outer
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { if recover() != nil { fmt.Println("outer") } }()
func() { panic("inner") }() }
func main() { run() }
