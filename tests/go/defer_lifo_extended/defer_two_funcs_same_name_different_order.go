// vybe-test: go/defer_lifo_extended/defer_two_funcs_same_name_different_order
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func a() { fmt.Println("a") }
func b() { fmt.Println("b") }
func main() { defer a()
defer b()
}
