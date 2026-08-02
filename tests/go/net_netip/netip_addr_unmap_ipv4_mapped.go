// vybe-test: go/net_netip/netip_addr_unmap_ipv4_mapped
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("::ffff:192.0.2.1")
_ = a.Unmap().String() }
