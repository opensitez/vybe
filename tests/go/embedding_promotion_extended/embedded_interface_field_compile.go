// vybe-test: go/embedding_promotion_extended/embedded_interface_field_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type speaker interface { talk() string }
type host struct { speaker }
func main() { _ = host{} }
