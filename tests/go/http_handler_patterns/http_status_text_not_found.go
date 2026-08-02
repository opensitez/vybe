// vybe-test: go/http_handler_patterns/http_status_text_not_found
// origin: languages/go/tests/go/test_http_handler_patterns.rs

package main
import "fmt"
import "net/http"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(http.StatusText(http.StatusNotFound)), "Not Found") }
