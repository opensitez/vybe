// vybe-test: go/net_textproto_compile/net_listen_tcp
// origin: languages/go/tests/go/test_net_textproto_compile.rs
// vybe-test-mode: compile

package main
import "net"
func main() { ln, _ := net.Listen("tcp", ":0")
if ln != nil { ln.Close() } }
