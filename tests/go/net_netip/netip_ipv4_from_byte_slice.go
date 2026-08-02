// vybe-test: go/net_netip/netip_ipv4_from_byte_slice
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { _, ok := netip.AddrFromSlice([]byte{1, 2, 3, 4})
_ = ok }
