// vybe-test: go/net_netip/netip_must_parse_addr_panics_on_bad
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { defer func() { _ = recover() }()
_ = netip.MustParseAddr("1.2.3.4") }
