// vybe-test: go/defer_panic_variants/recover_captures_bool_false_panic
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { value := recover()
fmt.Println(value == false) }()
panic(false) }
func main() { run() }
