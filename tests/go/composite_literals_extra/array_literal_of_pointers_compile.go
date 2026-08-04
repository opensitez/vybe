// vybe-test: go/composite_literals_extra/array_literal_of_pointers_compile
// origin: languages/go/tests/go/test_composite_literals_extra.rs
// vybe-test-mode: compile

package main
func main() { a, b := 1, 2
values := [2]*int{&a, &b}
_ = values }
