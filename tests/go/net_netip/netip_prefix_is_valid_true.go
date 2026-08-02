// vybe-test: go/net_netip/netip_prefix_is_valid_true
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { p, _ := netip.ParsePrefix("10.0.0.0/8")
_ = p.IsValid() }
