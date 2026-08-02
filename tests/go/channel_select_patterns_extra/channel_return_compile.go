// vybe-test: go/channel_select_patterns_extra/channel_return_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func build() chan int { return make(chan int, 1) }
func main() { _ = build() }
