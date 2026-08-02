// vybe-test: go/type_conversions_extra/method_on_named_type_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
type score int
func (s score) next() int { return int(s) + 1 }
func main() { _ = score(5).next() }
