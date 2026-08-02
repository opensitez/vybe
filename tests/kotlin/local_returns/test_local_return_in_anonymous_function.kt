// vybe-test: kotlin/local_returns/test_local_return_in_anonymous_function
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f: (Int) -> Int = fun(x: Int): Int {
                if (x == 0) return 5
                return x
            }
            __check((f(0)).toString(), "5")
            __check((f(2)).toString(), "2")
        }
