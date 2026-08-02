// vybe-test: kotlin/destructuring/test_nested_destructuring_on_return
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun nested(): Pair<Pair<Int, Int>, Int> = Pair(Pair(10, 20), 30)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (pair, tail) = nested()
            val (left, right) = pair
            __check((left + right + tail).toString(), "60")
        }
