// vybe-test: go/net_netip/netip_addr_is6_false_on_ipv4
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

func main() { a, _ := netip.ParseAddr("1.2.3.4")
__check(fmt.Sprint(a.Is6()), "false") }
