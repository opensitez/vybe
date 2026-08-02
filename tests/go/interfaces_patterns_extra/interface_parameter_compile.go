// vybe-test: go/interfaces_patterns_extra/interface_parameter_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
type reader interface { read() int }
func use(value reader) int { return value.read() }
func main() { _ = use }
