// vybe-test: go/interfaces_patterns_extra/interface_with_multiple_methods_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
type shape interface { area() int
perimeter() int }
func main() {}
