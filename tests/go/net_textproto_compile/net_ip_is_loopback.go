// vybe-test: go/net_textproto_compile/net_ip_is_loopback
// origin: languages/go/tests/go/test_net_textproto_compile.rs
// vybe-test-mode: compile

package main
import "net"
func main() { ip := net.ParseIP("127.0.0.1")
_ = ip.IsLoopback() }
