// vybe-test: go/methods_receivers_extra/method_on_alias_like_named_type_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type text string
func (t text) label() string { return string(t) }
func main() { var value text
_ = value.label() }
