// vybe-test: go/cover_net_extra/rpc_response_type
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
type Response = rpc.Response
func main() { var resp Response
_ = resp.ServiceMethod
_ = resp.Seq
_ = resp.Error }
