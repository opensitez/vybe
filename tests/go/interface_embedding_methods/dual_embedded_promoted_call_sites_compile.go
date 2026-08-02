// vybe-test: go/interface_embedding_methods/dual_embedded_promoted_call_sites_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type fetch interface { fetch() int }
type store interface { store(int) }
type cache interface { fetch
store }
type mem struct{}
func (mem) fetch() int { return 0 }
func (mem) store(int) {}
func main() { var c cache = mem{}
c.fetch()
c.store(1) }
