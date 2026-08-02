// vybe-test: go/net_textproto_compile/net_cidr_mask
// origin: languages/go/tests/go/test_net_textproto_compile.rs
// vybe-test-mode: compile

package main
import "net"
func main() { _ = net.CIDRMask(24, 32) }
