// vybe-test: kotlin/destructuring/test_destructuring_with_var_and_reassignment
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun makePair(): Pair<Int, Int> = Pair(4, 8)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var (x, y) = makePair()
            x += 1
            y += 2
            __check((x).toString(), "5")
            __check((y).toString(), "10")
        }
