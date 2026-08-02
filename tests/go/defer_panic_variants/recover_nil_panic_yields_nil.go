// vybe-test: go/defer_panic_variants/recover_nil_panic_yields_nil
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover() == nil) }()
panic(nil) }
func main() { run() }
