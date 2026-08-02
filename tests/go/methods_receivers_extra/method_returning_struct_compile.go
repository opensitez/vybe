// vybe-test: go/methods_receivers_extra/method_returning_struct_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int }
type builder struct{}
func (builder) build() point { return point{x: 1} }
func main() { _ = builder{}.build() }
