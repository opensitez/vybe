// vybe-test: go/stdlib_net_mail_index/suffixarray_lookup
// origin: languages/go/tests/go/test_stdlib_net_mail_index.rs
// vybe-test-mode: compile

package main
import "index/suffixarray"
func main() { idx := suffixarray.New([]byte("banana"))
_ = idx.Lookup([]byte("ana"), -1) }
