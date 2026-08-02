// vybe-test: go/cover_net_extra/rpc_serve_request_type
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
type Server = rpc.Server
func main() { var s Server
_ = s.ServeRequest }
