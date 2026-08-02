// vybe-test: go/interfaces_patterns_extra/interface_embedding_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
type reader interface { read() int }
type closer interface { close() }
type resource interface { reader
closer }
func main() {}
