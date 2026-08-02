// vybe-test: go/net_netip/netip_addr_port_parse_host_port
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

func main() { ap, _ := netip.ParseAddrPort("127.0.0.1:8080")
__check(fmt.Sprint(ap.String()), "127.0.0.1:8080") }
