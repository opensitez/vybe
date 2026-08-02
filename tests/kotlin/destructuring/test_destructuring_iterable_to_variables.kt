// vybe-test: kotlin/destructuring/test_destructuring_iterable_to_variables
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf("a", "b", "c")
            val (first, second, third) = source
            __check((first).toString(), "a")
            __check((second).toString(), "b")
            __check((third).toString(), "c")
        }
