// vybe-test: go/net_netip/netip_prefix_from_addr_and_bits
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("192.168.0.0")
_, _ = netip.PrefixFrom(a, 16) }
