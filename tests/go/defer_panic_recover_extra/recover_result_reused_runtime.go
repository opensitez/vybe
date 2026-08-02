// vybe-test: go/defer_panic_recover_extra/recover_result_reused_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func run() { defer func() { value := recover()
fmt.Println(value)
fmt.Println(value != nil) }()
panic(3) }
func main() { run() }
