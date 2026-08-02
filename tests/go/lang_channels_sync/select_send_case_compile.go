// vybe-test: go/lang_channels_sync/select_send_case_compile
// origin: languages/go/tests/go/test_lang_channels_sync.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
select { case ch <- 1: } }
