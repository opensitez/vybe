// vybe-test: go/defer_panic_recover_extra/defer_method_on_pointer_receiver_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
type counter struct { n int }
func (c *counter) show() { fmt.Println(c.n) }
func main() { value := &counter{n: 6}
defer value.show()
value.n = 9
}
