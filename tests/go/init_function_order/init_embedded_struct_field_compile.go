// vybe-test: go/init_function_order/init_embedded_struct_field_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
type base struct { x int }
type child struct { base
y int }
var c child
func init() { c.x = 1
c.y = 2 }
func main() { _ = c.x + c.y }
