// vybe-test: go/maphash_crypto_rand/maphash_string_sum
// origin: languages/go/tests/go/test_maphash_crypto_rand.rs

package main
import "fmt"
import "hash/maphash"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var h maphash.Hash
h.SetSeed(maphash.MakeSeed())
h.WriteString("vybe")
__check(fmt.Sprint(h.Sum64() > 0), "true") }
