// vybe-test: go/net_netip/netip_addr_as_slice_four_bytes
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("1.2.3.4")
_ = len(a.AsSlice()) }
