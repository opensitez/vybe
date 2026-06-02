use crate::helpers::*;

macro_rules! go_compile_test {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

macro_rules! go_run_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = run_prints($src);
            assert_eq!(out, $expected);
        }
    };
}

macro_rules! run_cases {
    ($( $name:ident => ($src:expr, $expected:expr), )*) => {
        $( go_run_test!($name, $src, $expected); )*
    };
}

macro_rules! compile_cases {
    ($( $name:ident => $src:expr, )*) => {
        $( go_compile_test!($name, $src); )*
    };
}

run_cases! {
    buffered_channel_cap_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int, 3); fmt.Println(cap(ch)); }", vec!["3"]),
    buffered_channel_len_initial_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int, 2); fmt.Println(len(ch)); }", vec!["0"]),
    buffered_channel_len_after_send_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int, 2); ch <- 1; fmt.Println(len(ch)); }", vec!["1"]),
    buffered_channel_receive_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 4; fmt.Println(<-ch); }", vec!["4"]),
    buffered_channel_fifo_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int, 2); ch <- 1; ch <- 2; fmt.Println(<-ch); fmt.Println(<-ch); }", vec!["1", "2"]),
    buffered_channel_len_after_receive_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 9; _ = <-ch; fmt.Println(len(ch)); }", vec!["0"]),
    channel_in_struct_field_cap_runtime => ("package main; import \"fmt\"; type holder struct { ch chan int }; func main() { value := holder{ch: make(chan int, 5)}; fmt.Println(cap(value.ch)); }", vec!["5"]),
    channel_pass_to_function_runtime => ("package main; import \"fmt\"; func fill(ch chan int) { ch <- 7 }; func main() { ch := make(chan int, 1); fill(ch); fmt.Println(<-ch); }", vec!["7"]),
    channel_return_from_function_runtime => ("package main; import \"fmt\"; func build() chan int { return make(chan int, 4) }; func main() { fmt.Println(cap(build())); }", vec!["4"]),
    select_default_runtime => ("package main; import \"fmt\"; func main() { select { default: fmt.Println(1) } }", vec!["1"]),
    select_receive_ready_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 8; select { case v := <-ch: fmt.Println(v); default: fmt.Println(0) } }", vec!["8"]),
    select_send_ready_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); select { case ch <- 5: fmt.Println(len(ch)); default: fmt.Println(0) } }", vec!["1"]),
    channel_make_zero_buffer_cap_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int); fmt.Println(cap(ch)); }", vec!["0"]),
    channel_two_sends_len_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int, 2); ch <- 3; ch <- 4; fmt.Println(len(ch)); }", vec!["2"]),
    channel_receive_after_two_sends_runtime => ("package main; import \"fmt\"; func main() { ch := make(chan int, 2); ch <- 6; ch <- 7; first := <-ch; second := <-ch; fmt.Println(first + second); }", vec!["13"]),
}

compile_cases! {
    channel_direction_send_only_compile => "package main; func sendData(ch chan<- int) { ch <- 1 }; func main() { _ = sendData }",
    channel_direction_recv_only_compile => "package main; func recvData(ch <-chan int) int { return <-ch }; func main() { _ = recvData }",
    close_buffered_channel_compile => "package main; func main() { ch := make(chan int, 1); close(ch) }",
    close_nil_channel_compile => "package main; func main() { var ch chan int; close(ch) }",
    select_multiple_cases_compile => "package main; func main() { ch1 := make(chan int, 1); ch2 := make(chan int, 1); select { case <-ch1: case ch2 <- 1: default: } }",
    select_assign_ok_compile => "package main; func main() { ch := make(chan int, 1); select { case v, ok := <-ch: _, _ = v, ok; default: } }",
    range_over_channel_compile => "package main; func main() { ch := make(chan int); go func() { close(ch) }(); for value := range ch { _ = value } }",
    goroutine_send_compile => "package main; func main() { ch := make(chan int); go func() { ch <- 1 }() }",
    goroutine_receive_compile => "package main; func main() { ch := make(chan int); go func() { _ = <-ch }() }",
    channel_in_slice_compile => "package main; func main() { _ = []chan int{make(chan int, 1)} }",
    channel_in_map_compile => "package main; func main() { _ = map[string]chan int{\"a\": make(chan int, 1)} }",
    channel_in_struct_compile => "package main; type holder struct { ch chan int }; func main() { _ = holder{ch: make(chan int, 1)} }",
    recv_only_param_compile => "package main; func use(ch <-chan int) { _, _ = <-ch }; func main() { _ = use }",
    send_only_param_compile => "package main; func use(ch chan<- int) { ch <- 1 }; func main() { _ = use }",
    channel_type_alias_compile => "package main; type numbers chan int; func main() { var ch numbers = make(chan int, 1); _ = ch }",
    select_empty_compile => "package main; func main() { select {} }",
    select_with_break_compile => "package main; func main() { ch := make(chan int, 1); select { case <-ch: break; default: } }",
    select_in_loop_compile => "package main; func main() { ch := make(chan int, 1); for { select { case <-ch: return; default: return } } }",
    nested_select_compile => "package main; func main() { ch := make(chan int, 1); select { default: select { default: } } }",
    channel_receive_two_value_compile => "package main; func main() { ch := make(chan int, 1); value, ok := <-ch; _, _ = value, ok }",
    channel_send_in_select_compile => "package main; func main() { ch := make(chan int, 1); select { case ch <- 1: default: } }",
    channel_receive_in_if_compile => "package main; func main() { ch := make(chan int, 1); if value := <-ch; value >= 0 { _ = value } }",
    channel_cap_builtin_compile => "package main; func main() { ch := make(chan int, 2); _ = cap(ch) }",
    channel_len_builtin_compile => "package main; func main() { ch := make(chan int, 2); _ = len(ch) }",
    channel_close_after_send_compile => "package main; func main() { ch := make(chan int, 1); ch <- 1; close(ch) }",
    select_recv_only_channel_compile => "package main; func main() { ch := make(chan int, 1); var recv <-chan int = ch; select { case <-recv: default: } }",
    select_send_only_channel_compile => "package main; func main() { ch := make(chan int, 1); var send chan<- int = ch; select { case send <- 1: default: } }",
    channel_return_compile => "package main; func build() chan int { return make(chan int, 1) }; func main() { _ = build() }",
    channel_param_return_compile => "package main; func passthrough(ch chan int) chan int { return ch }; func main() { _ = passthrough }",
    goroutine_closure_with_channel_compile => "package main; func main() { ch := make(chan int, 1); go func(local chan int) { local <- 1 }(ch) }",
    channel_make_in_var_decl_compile => "package main; var ch = make(chan int, 1); func main() { _ = ch }",
    channel_in_interface_compile => "package main; func main() { var value interface{} = make(chan int, 1); _ = value }",
    channel_compare_nil_compile => "package main; func main() { var ch chan int; _ = (ch == nil) }",
    channel_assign_variable_compile => "package main; func main() { first := make(chan int, 1); second := first; _ = second }",
    channel_select_named_case_compile => "package main; func main() { ch := make(chan int, 1); select { case value := <-ch: _ = value; default: } }",
}
