// vybe-test: go/interface_embedding_methods/nil_composite_promoted_read_call_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type reader interface { read() int }
type loader interface { reader }
func main() { var value loader
_ = value.read() }
