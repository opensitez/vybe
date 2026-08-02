// vybe-test: go/net_netip/netip_addr_port_bracketed_ipv6
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { _, _ = netip.ParseAddrPort("[2001:db8::1]:80") }
