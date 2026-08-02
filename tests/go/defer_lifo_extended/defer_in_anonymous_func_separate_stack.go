// vybe-test: go/defer_lifo_extended/defer_in_anonymous_func_separate_stack
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { defer fmt.Println("main")
func() { defer fmt.Println("anon")
}()
}
