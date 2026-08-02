// vybe-test: go/lang_builtins_control/channel_send_only_type
// origin: languages/go/tests/go/test_lang_builtins_control.rs
// vybe-test-mode: compile

package main
func main() { var ch chan<- int
_ = ch }
