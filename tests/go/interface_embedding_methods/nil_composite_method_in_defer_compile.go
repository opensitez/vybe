// vybe-test: go/interface_embedding_methods/nil_composite_method_in_defer_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type cleaner interface { clean() }
type janitor interface { cleaner }
func sweep(value janitor) { defer value.clean() }
func main() { sweep(nil) }
