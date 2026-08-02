// vybe-test: go/method_sets_pointer_value/named_interface_from_embedded_method_set_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type mover interface { move() int }
type legs struct{}
func (legs) move() int { return 1 }
type body struct { legs }
func main() { var m mover = body{}
_ = m }
