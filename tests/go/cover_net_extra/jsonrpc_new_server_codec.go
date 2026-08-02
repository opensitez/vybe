// vybe-test: go/cover_net_extra/jsonrpc_new_server_codec
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net"
import "net/rpc/jsonrpc"
func main() { conn, _ := net.Dial("tcp", "127.0.0.1:9999")
if conn != nil { defer conn.Close()
_ = jsonrpc.NewServerCodec(conn) } }
