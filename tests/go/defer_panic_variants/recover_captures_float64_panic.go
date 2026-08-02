// vybe-test: go/defer_panic_variants/recover_captures_float64_panic
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
panic(2.5) }
func main() { run() }
