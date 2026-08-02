// vybe-test: go/stdlib_net_mail_index/suffixarray_new
// origin: languages/go/tests/go/test_stdlib_net_mail_index.rs
// vybe-test-mode: compile

package main
import "index/suffixarray"
func main() { _ = suffixarray.New([]byte("banana")) }
