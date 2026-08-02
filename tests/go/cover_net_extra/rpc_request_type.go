// vybe-test: go/cover_net_extra/rpc_request_type
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
type Request = rpc.Request
func main() { var req Request
_ = req.ServiceMethod
_ = req.Seq
_ = req.Args }
