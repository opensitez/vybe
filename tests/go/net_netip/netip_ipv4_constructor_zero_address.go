// vybe-test: go/net_netip/netip_ipv4_constructor_zero_address
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { _ = netip.IPv4(0, 0, 0, 0).IsUnspecified() }
