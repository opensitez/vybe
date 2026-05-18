use crate::helpers::*;

#[test] fn method_on_custom_int() {
    let out = run_prints("package main; import \"fmt\"; type MyInt int; func (m MyInt) IsPositive() bool { return m > 0 }; func main() { var x MyInt = 5; fmt.Println(x.IsPositive()); }");
    assert_eq!(out, vec!["true"]);
}
#[test] fn method_on_custom_string() {
    let out = run_prints("package main; import \"fmt\"; type MyStr string; func (s MyStr) Len() int { return len(s) }; func main() { var s MyStr = \"hello\"; fmt.Println(s.Len()); }");
    assert_eq!(out, vec!["5"]);
}
#[test] fn method_pointer_receiver_modify() {
    let out = run_prints("package main; import \"fmt\"; type Counter struct { val int }; func (c *Counter) Add(n int) { c.val += n }; func main() { c := Counter{val: 0}; c.Add(5); fmt.Println(c.val); }");
    assert_eq!(out, vec!["5"]);
}
#[test] fn method_value_receiver_no_modify() {
    let out = run_prints("package main; import \"fmt\"; type Counter struct { val int }; func (c Counter) Add(n int) { c.val += n }; func main() { c := Counter{val: 0}; c.Add(5); fmt.Println(c.val); }");
    assert_eq!(out, vec!["0"]);
}
#[test] fn method_call_on_pointer() {
    let out = run_prints("package main; import \"fmt\"; type Box struct { size int }; func (b *Box) GetSize() int { return b.size }; func main() { b := &Box{size: 10}; fmt.Println(b.GetSize()); }");
    assert_eq!(out, vec!["10"]);
}
#[test] fn method_call_pointer_auto_deref() {
    let out = run_prints("package main; import \"fmt\"; type Box struct { size int }; func (b Box) GetSize() int { return b.size }; func main() { b := &Box{size: 10}; fmt.Println(b.GetSize()); }");
    assert_eq!(out, vec!["10"]);
}
#[test] fn method_call_value_auto_addr() {
    let out = run_prints("package main; import \"fmt\"; type Box struct { size int }; func (b *Box) SetSize(s int) { b.size = s }; func main() { b := Box{size: 0}; b.SetSize(10); fmt.Println(b.size); }");
    assert_eq!(out, vec!["10"]);
}
#[test] fn method_chaining() {
    let out = run_prints("package main; import \"fmt\"; type Calc struct { n int }; func (c *Calc) Add(x int) *Calc { c.n += x; return c }; func main() { c := &Calc{n: 0}; c.Add(5).Add(3); fmt.Println(c.n); }");
    assert_eq!(out, vec!["8"]);
}
#[test] fn method_same_name_different_types() {
    let out = run_prints("package main; import \"fmt\"; type A struct{}; func (A) Print() { fmt.Println(\"A\") }; type B struct{}; func (B) Print() { fmt.Println(\"B\") }; func main() { a := A{}; b := B{}; a.Print(); b.Print(); }");
    assert_eq!(out, vec!["A", "B"]);
}
#[test] fn method_expression() {
    let out = run_prints("package main; import \"fmt\"; type T struct { n int }; func (t T) Print() { fmt.Println(t.n) }; func main() { t := T{n: 42}; f := T.Print; f(t); }");
    assert_eq!(out, vec!["42"]);
}
