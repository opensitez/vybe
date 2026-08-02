// vybe-test: go/channel_select_patterns_extra/channel_in_map_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = map[string]chan int{"a": make(chan int, 1)} }
