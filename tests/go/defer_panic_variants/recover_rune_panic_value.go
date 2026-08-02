// vybe-test: go/defer_panic_variants/recover_rune_panic_value
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
panic(rune(65)) }
func main() { run() }
