// vybe-test: go/type_conversions_extra/named_type_conversion_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
type score int
func main() { _ = score(1) }
