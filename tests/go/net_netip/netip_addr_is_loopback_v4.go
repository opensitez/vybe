// vybe-test: go/net_netip/netip_addr_is_loopback_v4
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("127.0.0.1")
_ = a.IsLoopback() }
