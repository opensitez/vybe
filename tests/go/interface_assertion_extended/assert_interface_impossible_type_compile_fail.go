// vybe-test: go/interface_assertion_extended/assert_interface_impossible_type_compile_fail
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile-fail

package main
type reader interface { read() int }
type writer interface { write(int) }
func main() { var r reader
_ = r.(writer) }
