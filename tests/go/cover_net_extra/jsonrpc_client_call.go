// vybe-test: go/cover_net_extra/jsonrpc_client_call
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc/jsonrpc"
import "net"
func main() { conn, _ := net.Dial("tcp", "127.0.0.1:9999")
if conn != nil { defer conn.Close()
c := jsonrpc.NewClient(conn)
_ = c.Call("Arith.Add", struct{ A, B int }{1, 2}, new(int)) } }
