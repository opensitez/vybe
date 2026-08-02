// vybe-test: go/lang_channels_sync/select_assign_recv_compile
// origin: languages/go/tests/go/test_lang_channels_sync.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
ch <- 1
var v int
select { case v = <-ch: _ = v } }
