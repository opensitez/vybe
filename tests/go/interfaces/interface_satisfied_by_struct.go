// vybe-test: go/interfaces/interface_satisfied_by_struct
// origin: languages/go/tests/go/test_interfaces.rs
// vybe-test-mode: compile

package main
type Greeter interface { Greet() string } type Person struct { Name string } func (p Person) Greet() string { return "Hello " + p.Name } func main() { var g Greeter
g = Person{Name: "Alice"}
fmt.Println(g.Greet())
}
