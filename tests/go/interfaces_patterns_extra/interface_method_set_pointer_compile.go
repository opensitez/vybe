// vybe-test: go/interfaces_patterns_extra/interface_method_set_pointer_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
type reader interface { read() int }
type box struct{}
func (b *box) read() int { return 1 }
func main() { var value reader = &box{}
_ = value }
