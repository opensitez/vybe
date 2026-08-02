// vybe-test: kotlin/destructuring/test_destructuring_triple_and_computation
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun getTriple(): Triple<Int, Int, Int> = Triple(2, 4, 6)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (a, b, c) = getTriple()
            __check((a * b).toString(), "8")
            __check((c / b).toString(), "1")
        }
