// vybe-test: go/net_netip/netip_string_roundtrip_ipv4
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

func main() { orig := "203.0.113.5"
a, _ := netip.ParseAddr(orig)
__check(fmt.Sprint(a.String() == orig), "true") }
