// vybe-test: go/defer_lifo_extended/defer_order_with_three_named_funcs
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func p1() { fmt.Println(1) }
func p2() { fmt.Println(2) }
func p3() { fmt.Println(3) }
func main() { defer p1()
defer p2()
defer p3()
}
