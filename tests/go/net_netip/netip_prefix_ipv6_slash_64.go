// vybe-test: go/net_netip/netip_prefix_ipv6_slash_64
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { _, _ = netip.ParsePrefix("2001:db8::/64") }
