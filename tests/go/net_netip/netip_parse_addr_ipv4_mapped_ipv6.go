// vybe-test: go/net_netip/netip_parse_addr_ipv4_mapped_ipv6
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { _, _ = netip.ParseAddr("::ffff:192.0.2.1") }
