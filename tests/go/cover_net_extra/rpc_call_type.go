// vybe-test: go/cover_net_extra/rpc_call_type
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
type Call = rpc.Call
func main() { var call Call
_ = call.ServiceMethod
_ = call.Reply
_ = call.Error }
