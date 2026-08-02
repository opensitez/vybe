// vybe-test: go/net_netip/netip_addr_compare_equal
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("1.1.1.1")
b, _ := netip.ParseAddr("1.1.1.1")
_ = a.Compare(b) }
