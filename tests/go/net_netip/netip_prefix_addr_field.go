// vybe-test: go/net_netip/netip_prefix_addr_field
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { p, _ := netip.ParsePrefix("172.16.0.0/12")
_ = p.Addr().String() }
