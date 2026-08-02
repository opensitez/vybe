// vybe-test: go/net_netip/netip_addr_port_addr_field
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { ap, _ := netip.ParseAddrPort("203.0.113.7:9000")
_ = ap.Addr().String() }
