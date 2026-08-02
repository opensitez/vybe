// vybe-test: go/channel_buffered_patterns/buffered_fan_in_two_sources
// origin: languages/go/tests/go/test_channel_buffered_patterns.rs
// vybe-test-mode: compile

package main
func main() { out := make(chan int, 2)
a := make(chan int, 1)
b := make(chan int, 1)
go func() { out <- <-a }()
go func() { out <- <-b }() }
