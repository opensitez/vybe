// vybe-test: go/defer_lifo_extended/defer_mixed_func_and_literal_order
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func tag(s string) { fmt.Println(s) }
func main() { defer tag("one")
defer func() { fmt.Println("two") }()
defer tag("three")
}
