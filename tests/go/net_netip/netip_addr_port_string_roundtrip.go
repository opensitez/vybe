// vybe-test: go/net_netip/netip_addr_port_string_roundtrip
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { s := "198.51.100.2:22"
ap, _ := netip.ParseAddrPort(s)
_ = ap.String() == s }
