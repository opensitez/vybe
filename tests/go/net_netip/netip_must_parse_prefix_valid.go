// vybe-test: go/net_netip/netip_must_parse_prefix_valid
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { _ = netip.MustParsePrefix("0.0.0.0/0").Bits() }
