// vybe-test: go/nil_zero_semantics_extra/nil_channel_select_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { var ch chan int
select { case <-ch: default: } }
