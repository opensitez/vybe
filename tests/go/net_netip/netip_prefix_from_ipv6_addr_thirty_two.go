// vybe-test: go/net_netip/netip_prefix_from_ipv6_addr_thirty_two
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("2001:db8::")
p, _ := netip.PrefixFrom(a, 32)
_ = p.String() }
