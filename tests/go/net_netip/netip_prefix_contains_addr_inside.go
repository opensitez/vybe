// vybe-test: go/net_netip/netip_prefix_contains_addr_inside
// origin: languages/go/tests/go/test_net_netip.rs

package main
import "fmt"
import "net/netip"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p, _ := netip.ParsePrefix("10.0.0.0/8")
a, _ := netip.ParseAddr("10.1.2.3")
__check(fmt.Sprint(p.Contains(a)), "true") }
