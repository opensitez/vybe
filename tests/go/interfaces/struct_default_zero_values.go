// vybe-test: go/interfaces/struct_default_zero_values
// origin: languages/go/tests/go/test_interfaces.rs
// vybe-test-mode: compile

package main
type Counter struct { Count int } func main() { var c Counter
_ = c }
