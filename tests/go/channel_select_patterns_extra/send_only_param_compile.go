// vybe-test: go/channel_select_patterns_extra/send_only_param_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func use(ch chan<- int) { ch <- 1 }
func main() { _ = use }
