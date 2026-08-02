// vybe-test: go/init_function_order/init_mutates_struct_pointer_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
type node struct { next *node
val int }
var head node
func init() { head.val = 1
head.next = &head }
func main() { _ = head.next.val }
