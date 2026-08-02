// vybe-test: kotlin/destructuring/test_destructuring_pair_values
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun makePair(): Pair<Int, String> = Pair(1, "one")

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (id, label) = makePair()
            __check((id).toString(), "1")
            __check((label).toString(), "one")
        }
