// vybe-test: go/methods_receivers_extra/method_in_if_condition_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type gate struct { open bool }
func (g gate) ready() bool { return g.open }
func main() { value := gate{open: true}
if value.ready() { fmt.Println(1) } else { fmt.Println(0) } }
