// vybe-test: go/net_netip/netip_prefix_contains_prefix
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { outer, _ := netip.ParsePrefix("10.0.0.0/8")
inner, _ := netip.ParsePrefix("10.1.0.0/16")
_ = outer.ContainsPrefix(inner) }
