// vybe-test: go/interfaces_patterns_extra/interface_returning_interface_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
type reader interface { read() int }
type box struct{}
func (b box) read() int { return 1 }
func build() reader { return box{} }
func main() { _ = build() }
