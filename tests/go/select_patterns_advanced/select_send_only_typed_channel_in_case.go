// vybe-test: go/select_patterns_advanced/select_send_only_typed_channel_in_case
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
var out chan<- int = ch
select { case out <- 2: default: } }
