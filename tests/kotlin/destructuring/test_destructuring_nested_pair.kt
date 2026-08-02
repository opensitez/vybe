// vybe-test: kotlin/destructuring/test_destructuring_nested_pair
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun wrap(): Pair<Pair<Int, Int>, Int> = Pair(Pair(3, 4), 5)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (inner, tail) = wrap()
            val (first, second) = inner
            __check((first + second + tail).toString(), "12")
        }
