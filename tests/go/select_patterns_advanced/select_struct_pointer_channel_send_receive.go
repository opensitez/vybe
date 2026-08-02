// vybe-test: go/select_patterns_advanced/select_struct_pointer_channel_send_receive
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
type node struct { n int }
func main() { ch := make(chan *node, 1)
select { case ch <- &node{n: 8}: default: }
select { case <-ch: default: } }
