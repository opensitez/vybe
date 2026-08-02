// vybe-test: go/interface_nil_comparable/named_interface_typed_nil_parameter_compile
// origin: languages/go/tests/go/test_interface_nil_comparable.rs
// vybe-test-mode: compile

package main
type reader interface { read() int }
type book struct{}
func (b *book) read() int { return 1 }
func accept(value reader) bool { return value == nil }
func main() { var p *book
_ = accept(p) }
