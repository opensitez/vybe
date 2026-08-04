// vybe-test: go/types_advanced/unnamed_struct_compile
// origin: languages/go/tests/go/test_types_advanced.rs
// vybe-test-mode: compile

package main
func main() { s := struct{ Name string }{Name: "test"}
_ = s }
