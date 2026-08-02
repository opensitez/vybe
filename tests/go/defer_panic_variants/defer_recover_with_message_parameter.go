// vybe-test: go/defer_panic_variants/defer_recover_with_message_parameter
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func(label string) { if recover() != nil { fmt.Println(label) } }("handled")
panic("err") }
func main() { run() }
