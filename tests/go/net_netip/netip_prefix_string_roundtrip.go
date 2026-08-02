// vybe-test: go/net_netip/netip_prefix_string_roundtrip
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { s := "198.51.100.0/24"
p, _ := netip.ParsePrefix(s)
_ = p.String() == s }
