//! Channel close, range-over-channel, send/receive on buffered channels.

go_run_cases! {
    buffered_channel_len_cap => ("package main; import \"fmt\"; func main() { ch := make(chan int, 3); ch <- 1; ch <- 2; fmt.Println(len(ch)); fmt.Println(cap(ch)) }", vec!["2", "3"]),
    receive_from_buffered_fifo => ("package main; import \"fmt\"; func main() { ch := make(chan int, 2); ch <- 10; ch <- 20; fmt.Println(<-ch); fmt.Println(<-ch) }", vec!["10", "20"]),
    close_channel_range_sum => ("package main; import \"fmt\"; func main() { ch := make(chan int, 3); ch <- 1; ch <- 2; ch <- 3; close(ch); sum := 0; for v := range ch { sum += v }; fmt.Println(sum) }", vec!["6"]),
    close_channel_ok_false => ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 7; close(ch); v, ok := <-ch; fmt.Println(v); fmt.Println(ok) }", vec!["7", "true"]),
    receive_on_closed_zero_ok_false => ("package main; import \"fmt\"; func main() { ch := make(chan int); close(ch); v, ok := <-ch; fmt.Println(v); fmt.Println(ok) }", vec!["0", "false"]),
    nil_channel_send_blocks_compile_only => ("package main; import \"fmt\"; func main() { var ch chan int; fmt.Println(ch == nil) }", vec!["true"]),
}

go_compile_cases! {
    channel_direction_send_only => "package main; func send(ch chan<- int) { ch <- 1 }; func main() { ch := make(chan int, 1); send(ch) }",
    channel_direction_recv_only => "package main; func recv(ch <-chan int) int { return <-ch }; func main() { ch := make(chan int, 1); ch <- 2; _ = recv(ch) }",
    select_send_recv => "package main; func main() { ch := make(chan int, 1); select { case ch <- 1: default: } }",
    range_over_closed_empty_channel => "package main; func main() { ch := make(chan int); close(ch); for range ch { } }",
}
