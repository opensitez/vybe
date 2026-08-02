// vybe-test: go/net_netip/netip_ipv4_constructor_string
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

func main() { __check(fmt.Sprint(netip.IPv4(10, 0, 0, 1).String()), "10.0.0.1") }
