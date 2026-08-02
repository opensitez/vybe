// vybe-test: go/net_netip/netip_prefix_bits_returns_mask_length
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

func main() { p, _ := netip.ParsePrefix("192.0.2.0/26")
__check(fmt.Sprint(p.Bits()), "26") }
