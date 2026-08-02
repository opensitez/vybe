// vybe-test: go/channel_direction_extended/make_with_recv_only_type_compile
// origin: languages/go/tests/go/test_channel_direction_extended.rs
// vybe-test-mode: compile

package main
func main() { _ = make(<-chan int) }
