use crate::helpers::*;

#[test] fn interface_compile() {
    compile_ok("package main; type Animal interface { Sound() string } func main() {}");
}
#[test] fn interface_multiple_methods() {
    compile_ok("package main; type Shape interface { Area() float64; Perimeter() float64 } func main() {}");
}
#[test] fn interface_embedded() {
    compile_ok("package main; type Reader interface { Read() string } type Writer interface { Write(s string) } type ReadWriter interface { Reader; Writer } func main() {}");
}
#[test] fn interface_satisfied_by_struct() {
    compile_ok("package main; type Greeter interface { Greet() string } type Person struct { Name string } func (p Person) Greet() string { return \"Hello \" + p.Name } func main() { var g Greeter; g = Person{Name: \"Alice\"}; fmt.Println(g.Greet()); }");
}
#[test] fn interface_empty() {
    compile_ok("package main; func printAny(v interface{}) {} func main() { printAny(42); printAny(\"hello\"); }");
}
#[test] fn struct_with_interface_field() {
    compile_ok("package main; type Stringer interface { String() string } type Container struct { Value Stringer } func main() {}");
}
#[test] fn struct_method_value_receiver() {
    let out = run_prints("package main; import \"fmt\"; type Rect struct { W int; H int } func (r Rect) Area() int { return r.W * r.H } func main() { r := Rect{W: 3, H: 4}; fmt.Println(r.Area()); }");
    assert_eq!(out, vec!["12"]);
}
#[test] fn struct_method_perimeter() {
    let out = run_prints("package main; import \"fmt\"; type Rect struct { W int; H int } func (r Rect) Perimeter() int { return 2 * (r.W + r.H) } func main() { r := Rect{W: 3, H: 4}; fmt.Println(r.Perimeter()); }");
    assert_eq!(out, vec!["14"]);
}
#[test] fn struct_multiple_methods() {
    let out = run_prints("package main; import \"fmt\"; type Circle struct { R int } func (c Circle) Area() int { return 3 * c.R * c.R } func (c Circle) Diameter() int { return 2 * c.R } func main() { c := Circle{R: 5}; fmt.Println(c.Area()); fmt.Println(c.Diameter()); }");
    assert_eq!(out, vec!["75", "10"]);
}
#[test] fn struct_method_returns_string() {
    let out = run_prints("package main; import \"fmt\"; type Point struct { X int; Y int } func (p Point) String() string { return \"point\" } func main() { p := Point{X: 1, Y: 2}; fmt.Println(p.String()); }");
    assert_eq!(out, vec!["point"]);
}
#[test] fn struct_field_access() {
    let out = run_prints("package main; import \"fmt\"; type Dog struct { Name string; Age int } func main() { d := Dog{Name: \"Rex\", Age: 3}; fmt.Println(d.Name); fmt.Println(d.Age); }");
    assert_eq!(out, vec!["Rex", "3"]);
}
#[test] fn struct_default_zero_values() {
    compile_ok("package main; type Counter struct { Count int } func main() { var c Counter; _ = c }");
}
#[test] fn struct_nested() {
    let out = run_prints("package main; import \"fmt\"; type Address struct { City string } type Person struct { Name string; Addr Address } func main() { p := Person{Name: \"Bob\", Addr: Address{City: \"NY\"}}; fmt.Println(p.Addr.City); }");
    assert_eq!(out, vec!["NY"]);
}
#[test] fn struct_in_slice() {
    let out = run_prints("package main; import \"fmt\"; type Point struct { X int; Y int } func main() { pts := []Point{{X: 1, Y: 2}, {X: 3, Y: 4}}; fmt.Println(pts[0].X); fmt.Println(pts[1].Y); }");
    assert_eq!(out, vec!["1", "4"]);
}
#[test] fn struct_equality_check() {
    compile_ok("package main; type Vec struct { X int; Y int } func main() { v1 := Vec{X: 1, Y: 2}; v2 := Vec{X: 1, Y: 2}; _ = (v1 == v2) }");
}
#[test] fn interface_nil_check() {
    compile_ok("package main; type Doer interface { Do() }; func run(d Doer) { if d == nil { return }; d.Do() } func main() { run(nil) }");
}
#[test] fn method_chain_struct() {
    let out = run_prints("package main; import \"fmt\"; type Builder struct { Val int } func (b Builder) Add(n int) Builder { return Builder{Val: b.Val + n} } func main() { b := Builder{Val: 0}; b = b.Add(5); b = b.Add(3); fmt.Println(b.Val); }");
    assert_eq!(out, vec!["8"]);
}
#[test] fn struct_with_bool_field() {
    let out = run_prints("package main; import \"fmt\"; type Flag struct { Active bool } func main() { f := Flag{Active: true}; fmt.Println(f.Active); }");
    assert_eq!(out, vec!["true"]);
}
#[test] fn struct_method_modify_returns_new() {
    let out = run_prints("package main; import \"fmt\"; type Counter struct { N int } func (c Counter) Inc() Counter { return Counter{N: c.N + 1} } func main() { c := Counter{N: 0}; c = c.Inc(); c = c.Inc(); fmt.Println(c.N); }");
    assert_eq!(out, vec!["2"]);
}
#[test] fn type_alias_compile() {
    compile_ok("package main; type Celsius float64; type Fahrenheit float64; func main() { var c Celsius = 100; _ = c }");
}
#[test] fn struct_three_fields() {
    let out = run_prints("package main; import \"fmt\"; type Employee struct { Name string; Dept string; Salary int } func main() { e := Employee{Name: \"Alice\", Dept: \"Eng\", Salary: 90000}; fmt.Println(e.Name); fmt.Println(e.Dept); fmt.Println(e.Salary); }");
    assert_eq!(out, vec!["Alice", "Eng", "90000"]);
}
