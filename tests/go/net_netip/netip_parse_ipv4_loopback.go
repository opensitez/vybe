// vybe-test: go/net_netip/netip_parse_ipv4_loopback
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

func main() { a, _ := netip.ParseAddr("127.0.0.1")
__check(fmt.Sprint(a.String()), "127.0.0.1") }
