// vybe-test: go/net_netip/netip_ipv4_all_ones_broadcast
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

func main() { __check(fmt.Sprint(netip.IPv4(255, 255, 255, 255).String()), "255.255.255.255") }
