// vybe-test: kotlin/destructuring/test_destructuring_nullable_pair_with_default_fallback
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun maybe(): Pair<Int, Int>? = if (false) Pair(1, 2) else null

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (left, right) = maybe() ?: Pair(0, 0)
            __check((left + right).toString(), "0")
        }
