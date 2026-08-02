// vybe-test: go/interfaces_patterns_extra/named_interface_alias_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
type reader interface { read() int }
type alias = reader
func main() { var value alias
_ = value }
