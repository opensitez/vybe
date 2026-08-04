// vybe-test: go/interfaces/type_alias_compile
// origin: languages/go/tests/go/test_interfaces.rs
// vybe-test-mode: compile

package main
type Celsius float64
type Fahrenheit float64
func main() { var c Celsius = 100
_ = c }
