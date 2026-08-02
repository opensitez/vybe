// vybe-test: go/lang_generics_semantics/generic_chan_send_recv
// origin: languages/go/tests/go/test_lang_generics_semantics.rs
// vybe-test-mode: compile

package main
func Pump[T any](ch chan T, v T) { ch <- v }
func main() { ch := make(chan int,1)
Pump(ch, 1) }
