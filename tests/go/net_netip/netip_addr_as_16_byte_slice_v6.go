// vybe-test: go/net_netip/netip_addr_as_16_byte_slice_v6
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("::1")
_ = len(a.As16()) }
