// vybe-test: go/type_aliases/defined_int_pointer_receiver_compile
// origin: languages/go/tests/go/test_type_aliases.rs
// vybe-test-mode: compile

package main
type Counter int
func (c *Counter) bump() { *c = *c + 1 }
func main() { value := Counter(0)
value.bump()
_ = value }
