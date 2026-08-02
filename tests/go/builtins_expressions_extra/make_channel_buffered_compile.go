// vybe-test: go/builtins_expressions_extra/make_channel_buffered_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan string, 3)
_ = ch }
