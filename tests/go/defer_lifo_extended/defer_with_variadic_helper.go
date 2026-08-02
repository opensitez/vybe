// vybe-test: go/defer_lifo_extended/defer_with_variadic_helper
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func show(parts ...string) { fmt.Println(len(parts)) }
func main() { defer show("a", "b")
}
