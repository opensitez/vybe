// vybe-test: go/net_textproto_compile/net_resolve_tcp_addr
// origin: languages/go/tests/go/test_net_textproto_compile.rs
// vybe-test-mode: compile

package main
import "net"
func main() { _, _ = net.ResolveTCPAddr("tcp", ":80") }
