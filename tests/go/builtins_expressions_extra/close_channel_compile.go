// vybe-test: go/builtins_expressions_extra/close_channel_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
close(ch) }
