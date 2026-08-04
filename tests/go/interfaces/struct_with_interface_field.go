// vybe-test: go/interfaces/struct_with_interface_field
// origin: languages/go/tests/go/test_interfaces.rs
// vybe-test-mode: compile

package main
type Stringer interface { String() string } type Container struct { Value Stringer } func main() {}
