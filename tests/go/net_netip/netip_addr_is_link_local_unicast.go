// vybe-test: go/net_netip/netip_addr_is_link_local_unicast
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("fe80::1")
_ = a.IsLinkLocalUnicast() }
