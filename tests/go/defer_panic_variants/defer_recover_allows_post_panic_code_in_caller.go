// vybe-test: go/defer_panic_variants/defer_recover_allows_post_panic_code_in_caller
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { recover() }()
panic("x")
fmt.Println("skip") }
func main() { run()
fmt.Println("after") }
