// vybe-test: go/interface_nil_comparable/nil_channel_compare_nil_compile
// origin: languages/go/tests/go/test_interface_nil_comparable.rs
// vybe-test-mode: compile

package main
func main() { var ch chan int
_ = (ch == nil) }
