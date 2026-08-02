// vybe-test: kotlin/named_arguments/test_named_arguments_named_call_for_unary_like
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun score(base: Int, delta: Int = 1): Int = base + delta
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(base = 9)).toString(), "10")
            __check((score(base = 9, delta = 0)).toString(), "9")
        }
