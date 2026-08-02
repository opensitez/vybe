// vybe-test: go/defer_lifo_extended/defer_lifo_with_panic_middle_frame
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer fmt.Println("c")
defer fmt.Println("b")
defer func() { recover() }()
fmt.Println("a")
panic("x") }
func main() { run() }
