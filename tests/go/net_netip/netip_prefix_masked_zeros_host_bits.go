// vybe-test: go/net_netip/netip_prefix_masked_zeros_host_bits
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

func main() { p, _ := netip.ParsePrefix("192.168.1.129/24")
__check(fmt.Sprint(p.Masked().String()), "192.168.1.0/24") }
