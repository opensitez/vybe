// vybe-test: go/net_netip/netip_addr_port_from_components
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("192.0.2.1")
_ = netip.AddrPortFrom(a, 53).String() }
