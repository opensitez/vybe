// vybe-test: go/defer_lifo_extended/defer_with_recover_then_next_defer
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer fmt.Println("last")
defer func() { recover() }()
defer fmt.Println("mid")
panic("p") }
func main() { run() }
