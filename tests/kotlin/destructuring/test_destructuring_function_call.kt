// vybe-test: kotlin/destructuring/test_destructuring_function_call
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun splitValue(): Pair<String, Int> = Pair("value", 9)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (text, count) = splitValue()
            __check((text + count.toString()).toString(), "value9")
        }
