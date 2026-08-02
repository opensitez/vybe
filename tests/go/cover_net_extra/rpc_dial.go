// vybe-test: go/cover_net_extra/rpc_dial
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
func main() { _, _ = rpc.Dial("tcp", "127.0.0.1:9999") }
