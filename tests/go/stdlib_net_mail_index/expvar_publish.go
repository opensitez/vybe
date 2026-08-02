// vybe-test: go/stdlib_net_mail_index/expvar_publish
// origin: languages/go/tests/go/test_stdlib_net_mail_index.rs
// vybe-test-mode: compile

package main
import "expvar"
func main() { expvar.Publish("hits", new(expvar.Int)) }
