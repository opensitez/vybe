// vybe-test: go/net_netip/netip_parse_addr_port_missing_port
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { _, err := netip.ParseAddrPort("127.0.0.1")
_ = err != nil }
