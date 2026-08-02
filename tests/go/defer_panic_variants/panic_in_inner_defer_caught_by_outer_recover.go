// vybe-test: go/defer_panic_variants/panic_in_inner_defer_caught_by_outer_recover
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { if recover() != nil { fmt.Println("caught") } }()
defer func() { panic("inner") }() }
func main() { run() }
