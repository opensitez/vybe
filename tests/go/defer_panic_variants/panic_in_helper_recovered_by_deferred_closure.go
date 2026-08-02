// vybe-test: go/defer_panic_variants/panic_in_helper_recovered_by_deferred_closure
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func boom() { panic("fail") }
func run() { defer func() { if recover() != nil { fmt.Println("saved") } }()
boom() }
func main() { run() }
