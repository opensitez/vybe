// vybe-test: go/net_netip/netip_addr_with_zone_id
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("fe80::1%eth0")
_ = a.WithZone("eth0").Zone() }
