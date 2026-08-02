// vybe-test: go/interfaces_patterns_extra/interface_in_slice_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
type reader interface { read() int }
func main() { var values []reader
_ = values }
