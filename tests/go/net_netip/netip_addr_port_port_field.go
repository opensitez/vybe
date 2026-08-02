// vybe-test: go/net_netip/netip_addr_port_port_field
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

func main() { ap, _ := netip.ParseAddrPort("[::1]:443")
__check(fmt.Sprint(ap.Port()), "443") }
