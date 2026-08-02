// vybe-test: kotlin/destructuring/test_destructuring_with_triple
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val triple = Triple("x", 5, true)
            val (name, count, active) = triple
            __check((name).toString(), "x")
            __check((count).toString(), "5")
            __check((active).toString(), "true")
        }
