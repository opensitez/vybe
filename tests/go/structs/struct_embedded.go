// vybe-test: go/structs/struct_embedded
// origin: languages/go/tests/go/test_structs.rs
// vybe-test-mode: compile

package main
type Base struct { ID int } type Child struct { Base
Name string } func main() { c := Child{Name: "test"}
_ = c }
