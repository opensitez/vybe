// vybe-test: go/cover_net_extra/rpc_server_handle_http
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
func main() { s := rpc.NewServer()
s.HandleHTTP() }
