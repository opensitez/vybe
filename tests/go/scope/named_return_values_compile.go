// vybe-test: go/scope/named_return_values_compile
// origin: languages/go/tests/go/test_scope.rs
// vybe-test-mode: compile

package main
func divide(a, b float64) (result float64, err bool) { if b == 0 { err = true
return }
result = a / b
return } func main() { _, _ = divide(10, 2) }
