// vybe-test: go/fmt_errors_print/fscanf_bytes_reader_float
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var f float64
c, _ := fmt.Fscanf(bytes.NewReader([]byte("2.5")), "%f", &f)
__check(fmt.Sprint(c) + " " + fmt.Sprint(f), "1 2.5") }
