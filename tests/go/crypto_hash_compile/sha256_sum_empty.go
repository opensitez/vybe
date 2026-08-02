// vybe-test: go/crypto_hash_compile/sha256_sum_empty
// origin: languages/go/tests/go/test_crypto_hash_compile.rs

package main
import "fmt"
import "crypto/sha256"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { sum := sha256.Sum256([]byte{})
__check(fmt.Sprint(len(sum)), "32") }
