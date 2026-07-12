use crate::helpers::*;

#[test]
fn go_routine_call() {
    compile_ok("package main; func doWork() {}; func main() { go doWork(); }");
}
#[test]
fn go_routine_closure() {
    compile_ok("package main; func main() { go func() { }() }");
}
#[test]
fn go_routine_method() {
    compile_ok(
        "package main; type Worker struct{}; func (w Worker) Work() {}; func main() { w := Worker{}; go w.Work(); }",
    );
}
#[test]
fn channel_make() {
    compile_ok("package main; func main() { ch := make(chan int); _ = ch }");
}
#[test]
fn channel_make_buffered() {
    compile_ok("package main; func main() { ch := make(chan int, 10); _ = ch }");
}
#[test]
fn channel_send() {
    compile_ok("package main; func main() { ch := make(chan int); go func() { ch <- 1 }() }");
}
#[test]
fn channel_receive() {
    compile_ok("package main; func main() { ch := make(chan int); go func() { <-ch }() }");
}
#[test]
fn channel_receive_assign() {
    compile_ok(
        "package main; func main() { ch := make(chan int); go func() { v := <-ch; _ = v }() }",
    );
}
#[test]
fn channel_receive_ok_idiom() {
    compile_ok(
        "package main; func main() { ch := make(chan int); go func() { v, ok := <-ch; _, _ = v, ok }() }",
    );
}
#[test]
fn channel_close() {
    compile_ok("package main; func main() { ch := make(chan int); close(ch); }");
}
#[test]
fn channel_range() {
    compile_ok(
        "package main; func main() { ch := make(chan int); go func() { for v := range ch { _ = v } }() }",
    );
}
#[test]
fn channel_type_send_only() {
    compile_ok("package main; func sendData(ch chan<- int) { ch <- 1 }; func main() {}");
}
#[test]
fn channel_type_recv_only() {
    compile_ok("package main; func recvData(ch <-chan int) { <-ch }; func main() {}");
}
#[test]
fn select_empty() {
    compile_ok("package main; func main() { select {} }");
}
#[test]
fn select_cases() {
    compile_ok(
        "package main; func main() { ch1 := make(chan int); ch2 := make(chan int); select { case <-ch1: case ch2 <- 1: default: } }",
    );
}
#[test]
fn select_assign() {
    compile_ok(
        "package main; func main() { ch := make(chan int); select { case v := <-ch: _ = v; default: } }",
    );
}
#[test]
fn select_ok_idiom() {
    compile_ok(
        "package main; func main() { ch := make(chan int); select { case v, ok := <-ch: _, _ = v, ok; default: } }",
    );
}
