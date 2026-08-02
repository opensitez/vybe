// vybe-test: go/net_netip/netip_prefix_overlaps_disjoint
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParsePrefix("10.0.0.0/8")
b, _ := netip.ParsePrefix("192.168.0.0/16")
_ = a.Overlaps(b) }
