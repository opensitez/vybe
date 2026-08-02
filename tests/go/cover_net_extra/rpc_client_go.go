// vybe-test: go/cover_net_extra/rpc_client_go
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
type Args struct { A, B int }
func main() { c, _ := rpc.Dial("tcp", "127.0.0.1:9999")
if c != nil { defer c.Close()
_ = c.Go("Arith.Add", &Args{1, 2}, new(int), make(chan *rpc.Call, 1)) } }
