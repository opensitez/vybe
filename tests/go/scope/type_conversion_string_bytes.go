// vybe-test: go/scope/type_conversion_string_bytes
// origin: languages/go/tests/go/test_scope.rs
// vybe-test-mode: compile

package main
func main() { s := "hello"
b := []byte(s)
_ = b }
