// vybe-test: go/net_netip/netip_ipv6_from_sixteen_byte_slice
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { _, ok := netip.AddrFromSlice(make([]byte, 16))
_ = ok }
