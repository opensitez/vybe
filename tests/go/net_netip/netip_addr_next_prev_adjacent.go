// vybe-test: go/net_netip/netip_addr_next_prev_adjacent
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("192.0.2.1")
_ = a.Next().String()
_ = a.Prev().String() }
