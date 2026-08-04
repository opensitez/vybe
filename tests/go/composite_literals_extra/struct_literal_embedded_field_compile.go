// vybe-test: go/composite_literals_extra/struct_literal_embedded_field_compile
// origin: languages/go/tests/go/test_composite_literals_extra.rs
// vybe-test-mode: compile

package main
type inner struct { value int }
type outer struct { inner }
func main() { _ = outer{inner: inner{value: 3}} }
