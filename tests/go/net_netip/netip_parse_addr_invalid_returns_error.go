// vybe-test: go/net_netip/netip_parse_addr_invalid_returns_error
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { _, err := netip.ParseAddr("not-an-ip")
_ = err != nil }
