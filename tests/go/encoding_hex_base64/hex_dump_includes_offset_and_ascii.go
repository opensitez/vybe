// vybe-test: go/encoding_hex_base64/hex_dump_includes_offset_and_ascii
// origin: languages/go/tests/go/test_encoding_hex_base64.rs

package main
import "fmt"
import "encoding/hex"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { dump := string(hex.Dump([]byte("ab")))
__check(fmt.Sprint(len(dump) > 0), "true")
__check(fmt.Sprint(dump[0:8]), "00000000") }
