// vybe-test: go/interfaces/struct_equality_check
// origin: languages/go/tests/go/test_interfaces.rs
// vybe-test-mode: compile

package main
type Vec struct { X int
Y int } func main() { v1 := Vec{X: 1, Y: 2}
v2 := Vec{X: 1, Y: 2}
_ = (v1 == v2) }
