// vybe-test: go/net_netip/netip_addr_less_than_ordering
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("1.0.0.1")
b, _ := netip.ParseAddr("1.0.0.2")
_ = a.Less(b) }
