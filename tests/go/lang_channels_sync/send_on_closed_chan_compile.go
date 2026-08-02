// vybe-test: go/lang_channels_sync/send_on_closed_chan_compile
// origin: languages/go/tests/go/test_lang_channels_sync.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
close(ch)
ch <- 1 }
