// vybe-test: go/net_netip/netip_string_roundtrip_ipv6
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

func main() { orig := "2001:db8:85a3::8a2e:370:7334"
a, _ := netip.ParseAddr(orig)
__check(fmt.Sprint(a.String() == orig), "true") }
