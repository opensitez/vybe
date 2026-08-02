// vybe-test: go/defer_panic_variants/defer_lifo_interleaved_with_work_before_panic
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer fmt.Println("third")
defer fmt.Println("second")
defer func() { recover() }()
fmt.Println("first")
panic("stop") }
func main() { run() }
