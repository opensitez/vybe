// vybe-test: kotlin/tuples/test_triple_destructuring_skips_first_component
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (_, mid, last) = Triple("a", 14, 15)
            __check((mid).toString(), "14")
            __check((last).toString(), "15")
        }
