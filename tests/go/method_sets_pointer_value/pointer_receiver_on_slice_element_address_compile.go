// vybe-test: go/method_sets_pointer_value/pointer_receiver_on_slice_element_address_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type item struct { n int }
func (i *item) bump() { i.n++ }
func main() { items := []item{{n: 1}}
items[0].bump() }
