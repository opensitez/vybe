// vybe-test: go/defer_panic_variants/recover_type_switch_on_int_panic
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { switch value := recover().(type) { case int: fmt.Println(value + 1)
default: fmt.Println(0) } }()
panic(6) }
func main() { run() }
