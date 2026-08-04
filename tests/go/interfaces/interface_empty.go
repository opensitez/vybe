// vybe-test: go/interfaces/interface_empty
// origin: languages/go/tests/go/test_interfaces.rs
// vybe-test-mode: compile

package main
func printAny(v interface{}) {} func main() { printAny(42)
printAny("hello")
}
