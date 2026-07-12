use crate::helpers::*;

#[test]
fn struct_declaration() {
    compile_ok("package main; type Person struct { Name string; Age int } func main() {}");
}
#[test]
fn struct_literal() {
    let out = run_prints(
        "package main; import \"fmt\"; type Person struct { Name string; Age int } func main() { p := Person{Name: \"Alice\", Age: 30}; fmt.Println(p.Name); fmt.Println(p.Age); }",
    );
    assert_eq!(out, vec!["Alice", "30"]);
}
#[test]
fn struct_method() {
    let out = run_prints(
        "package main; import \"fmt\"; type Person struct { Name string; Age int } func (p Person) Greet() { fmt.Println(p.Name); } func main() { p := Person{Name: \"Bob\", Age: 25}; p.Greet(); }",
    );
    assert_eq!(out, vec!["Bob"]);
}
#[test]
fn interface_declaration() {
    compile_ok("package main; type Greeter interface { Greet() } func main() {}");
}
#[test]
fn struct_update_field() {
    let out = run_prints(
        "package main; import \"fmt\"; type Counter struct { N int } func main() { c := Counter{N: 0}; c.N = 5; fmt.Println(c.N); }",
    );
    assert_eq!(out, vec!["5"]);
}
#[test]
fn struct_two_string_fields() {
    let out = run_prints(
        "package main; import \"fmt\"; type Point struct { X int; Y int } func main() { p := Point{X: 3, Y: 4}; fmt.Println(p.X); fmt.Println(p.Y); }",
    );
    assert_eq!(out, vec!["3", "4"]);
}
#[test]
fn struct_method_returns_int() {
    let out = run_prints(
        "package main; import \"fmt\"; type Rect struct { W int; H int } func (r Rect) Area() int { return r.W * r.H } func main() { r := Rect{W: 5, H: 6}; fmt.Println(r.Area()); }",
    );
    assert_eq!(out, vec!["30"]);
}
#[test]
fn struct_method_uses_self_field() {
    let out = run_prints(
        "package main; import \"fmt\"; type Box struct { Side int } func (b Box) Volume() int { return b.Side * b.Side * b.Side } func main() { bx := Box{Side: 3}; fmt.Println(bx.Volume()); }",
    );
    assert_eq!(out, vec!["27"]);
}
#[test]
fn struct_literal_partial() {
    compile_ok(
        "package main; type Config struct { Host string; Port int; Debug bool } func main() { c := Config{Host: \"localhost\"}; _ = c }",
    );
}
#[test]
fn struct_in_func_param() {
    let out = run_prints(
        "package main; import \"fmt\"; type Vec struct { X int; Y int } func dotProduct(a Vec, b Vec) int { return a.X*b.X + a.Y*b.Y } func main() { v1 := Vec{X: 2, Y: 3}; v2 := Vec{X: 4, Y: 5}; fmt.Println(dotProduct(v1, v2)); }",
    );
    assert_eq!(out, vec!["23"]);
}
#[test]
fn struct_slice() {
    let out = run_prints(
        "package main; import \"fmt\"; type Item struct { Val int } func main() { items := []Item{{Val: 1}, {Val: 2}, {Val: 3}}; total := 0; for _, it := range items { total = total + it.Val }; fmt.Println(total); }",
    );
    assert_eq!(out, vec!["6"]);
}
#[test]
fn struct_embedded() {
    compile_ok(
        "package main; type Base struct { ID int } type Child struct { Base; Name string } func main() { c := Child{Name: \"test\"}; _ = c }",
    );
}
#[test]
fn struct_with_slice_field() {
    let out = run_prints(
        "package main; import \"fmt\"; type List struct { Items []int } func main() { l := List{Items: []int{1, 2, 3}}; fmt.Println(len(l.Items)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn struct_with_map_field() {
    let out = run_prints(
        "package main; import \"fmt\"; type Registry struct { Data map[string]int } func main() { r := Registry{Data: map[string]int{\"a\": 1}}; fmt.Println(r.Data[\"a\"]); }",
    );
    assert_eq!(out, vec!["1"]);
}
#[test]
fn method_with_params() {
    let out = run_prints(
        "package main; import \"fmt\"; type Wallet struct { Balance int } func (w Wallet) Withdraw(amount int) int { return w.Balance - amount } func main() { w := Wallet{Balance: 100}; fmt.Println(w.Withdraw(30)); }",
    );
    assert_eq!(out, vec!["70"]);
}
#[test]
fn method_bool_return() {
    let out = run_prints(
        "package main; import \"fmt\"; type Validator struct { Min int; Max int } func (v Validator) InRange(n int) bool { return n >= v.Min && n <= v.Max } func main() { val := Validator{Min: 1, Max: 10}; fmt.Println(val.InRange(5)); fmt.Println(val.InRange(15)); }",
    );
    assert_eq!(out, vec!["true", "false"]);
}
#[test]
fn two_structs_interact() {
    let out = run_prints(
        "package main; import \"fmt\"; type A struct { Val int } type B struct { Val int } func combine(a A, b B) int { return a.Val + b.Val } func main() { fmt.Println(combine(A{Val: 3}, B{Val: 7})); }",
    );
    assert_eq!(out, vec!["10"]);
}
#[test]
fn struct_method_string_concat() {
    let out = run_prints(
        "package main; import \"fmt\"; type Name struct { First string; Last string } func (n Name) Full() string { return n.First + \" \" + n.Last } func main() { name := Name{First: \"John\", Last: \"Doe\"}; fmt.Println(name.Full()); }",
    );
    assert_eq!(out, vec!["John Doe"]);
}
