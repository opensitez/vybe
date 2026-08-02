// vybe-test: go/lang_channels_sync/double_close_compile
// origin: languages/go/tests/go/test_lang_channels_sync.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
close(ch)
close(ch) }
