// vybe-test: go/methods_receivers_extra/method_on_named_type_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type score int
func (s score) next() int { return int(s) + 1 }
func main() { var value score
_ = value.next() }
