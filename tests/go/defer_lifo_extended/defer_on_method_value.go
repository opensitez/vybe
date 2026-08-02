// vybe-test: go/defer_lifo_extended/defer_on_method_value
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
type counter struct { n int }
func (c counter) val() { fmt.Println(c.n) }
func main() { c := counter{n: 5}
defer c.val()
c.n = 9
}
