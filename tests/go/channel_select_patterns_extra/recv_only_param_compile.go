// vybe-test: go/channel_select_patterns_extra/recv_only_param_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func use(ch <-chan int) { _, _ = <-ch }
func main() { _ = use }
