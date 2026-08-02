// vybe-test: go/defer_panic_recover_extra/recover_type_preserved_string_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func run() { defer func() { value := recover()
fmt.Println(value == "err") }()
panic("err") }
func main() { run() }
