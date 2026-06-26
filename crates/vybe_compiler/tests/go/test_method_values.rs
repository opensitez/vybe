//! Method values and method expressions — distinct call forms.

use crate::helpers::*;

go_run_cases! {
    method_value_on_literal => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) twice() int { return c.n * 2 }; func main() { f := counter{n:3}.twice; fmt.Println(f()) }", vec!["6"]),
    method_expression_requires_receiver => ("package main; import \"fmt\"; type box struct { v int }; func (b box) get() int { return b.v }; func main() { f := box.get; fmt.Println(f(box{v:9})) }", vec!["9"]),
    pointer_receiver_method_value => ("package main; import \"fmt\"; type acc struct { sum int }; func (a *acc) add(x int) { a.sum += x }; func main() { a := &acc{}; inc := a.add; inc(4); inc(5); fmt.Println(a.sum) }", vec!["9"]),
    interface_method_value => ("package main; import \"fmt\"; type greeter interface { greet() string }; type hi struct{}; func (h hi) greet() string { return \"hi\" }; func main() { var g greeter = hi{}; f := g.greet; fmt.Println(f()) }", vec!["hi"]),
}

go_compile_cases! {
    method_value_stored_in_struct_field => "package main; type fn func() int; type holder struct { call fn }; func (c counter) val() int { return 1 }; type counter struct{}; func main() { _ = holder{} }",
    method_expression_with_pointer_receiver => "package main; type T struct{}; func (t *T) M() {}; func main() { f := (*T).M; _ = f }",
    method_on_function_type => "package main; type F func(); func (f F) call() { f() }; func main() { var fn F = func() {}; _ = fn.call }",
}
