// vybe-test: go/defer_lifo_extended/defer_lifo_preserves_registration_order_on_panic
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer fmt.Println(1)
defer fmt.Println(2)
defer func() { recover() }()
panic("x") }
func main() { run() }
