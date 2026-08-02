// vybe-test: go/defer_panic_variants/defer_recover_prints_panic_message
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { if r := recover(); r != nil { fmt.Println(r) } }()
panic("halt") }
func main() { run() }
