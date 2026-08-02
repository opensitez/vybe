// vybe-test: go/interfaces_patterns_extra/interface_in_struct_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
type reader interface { read() int }
type holder struct { value reader }
func main() { _ = holder{} }
