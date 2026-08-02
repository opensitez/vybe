// vybe-test: go/net_textproto_compile/net_parse_cidr
// origin: languages/go/tests/go/test_net_textproto_compile.rs
// vybe-test-mode: compile

package main
import "net"
func main() { _, _, _ = net.ParseCIDR("10.0.0.0/8") }
