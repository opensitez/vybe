// vybe-test: go/cover_net_extra/jsonrpc_serve_conn
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net"
import "net/rpc/jsonrpc"
func main() { ln, _ := net.Listen("tcp", "127.0.0.1:0")
defer ln.Close()
conn, _ := ln.Accept()
if conn != nil { jsonrpc.ServeConn(conn) } }
