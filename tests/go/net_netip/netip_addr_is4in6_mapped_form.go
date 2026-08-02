// vybe-test: go/net_netip/netip_addr_is4in6_mapped_form
// origin: languages/go/tests/go/test_net_netip.rs
// vybe-test-mode: compile

package main
import "net/netip"
func main() { a, _ := netip.ParseAddr("::ffff:1.2.3.4")
_ = a.Is4In6() }
