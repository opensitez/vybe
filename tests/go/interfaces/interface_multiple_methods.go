// vybe-test: go/interfaces/interface_multiple_methods
// origin: languages/go/tests/go/test_interfaces.rs
// vybe-test-mode: compile

package main
type Shape interface { Area() float64
Perimeter() float64 } func main() {}
