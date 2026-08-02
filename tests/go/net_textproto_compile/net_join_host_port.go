// vybe-test: go/net_textproto_compile/net_join_host_port
// origin: languages/go/tests/go/test_net_textproto_compile.rs
// vybe-test-mode: compile

package main
import "net"
func main() { _, _ = net.JoinHostPort("127.0.0.1", "80") }
