// vybe-test: go/interface_embedding_methods/overlapping_method_three_interfaces_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type a interface { ping() int }
type b interface { ping() int }
type c interface { ping() int }
type trio interface { a
b
c }
type echo struct{}
func (echo) ping() int { return 1 }
func main() { var value trio = echo{}
_ = value.ping() }
