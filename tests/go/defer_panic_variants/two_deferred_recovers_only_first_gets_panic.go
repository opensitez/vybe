// vybe-test: go/defer_panic_variants/two_deferred_recovers_only_first_gets_panic
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover() == nil) }()
defer func() { fmt.Println(recover() != nil) }()
panic("boom") }
func main() { run() }
