// vybe-test: go/cover_net_extra/rpc_default_server_handle_http
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
func main() { rpc.DefaultServer.HandleHTTP() }
