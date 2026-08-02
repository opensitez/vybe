// vybe-test: go/cover_net_extra/rpc_server_accept
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net"
import "net/rpc"
func main() { s := rpc.NewServer()
ln, _ := net.Listen("tcp", "127.0.0.1:0")
defer ln.Close()
go s.Accept(ln) }
