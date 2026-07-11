//! Buffered channels — one smoke test per distinct API (cap, len, blocking, fan-in).

go_run_cases! {
    buffered_cap => ("package main; import \"fmt\"; func main() { ch := make(chan int, 4); fmt.Println(cap(ch)) }", vec!["4"]),
    buffered_len_after_send => ("package main; import \"fmt\"; func main() { ch := make(chan int, 2); ch <- 1; fmt.Println(len(ch)) }", vec!["1"]),
}

go_compile_cases! {
    buffered_send_blocks_when_full => "package main; func main() { ch := make(chan int, 1); ch <- 1; ch <- 2 }",
    buffered_recv_blocks_when_empty => "package main; func main() { ch := make(chan int, 1); _ = <-ch }",
    buffered_fan_in_two_sources => "package main; func main() { out := make(chan int, 2); a := make(chan int, 1); b := make(chan int, 1); go func() { out <- <-a }(); go func() { out <- <-b }() }",
}
