// vybe-test: go/net_netip/netip_parse_addr_unspecified_ipv6
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("::")
_ = a.IsUnspecified() }
