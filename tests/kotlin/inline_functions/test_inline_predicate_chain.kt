// vybe-test: kotlin/inline_functions/test_inline_predicate_chain
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun check(value: Int, tests: (Int) -> Boolean, onFail: () -> String): String {
            return if (tests(value)) "ok" else onFail()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((check(5, { it > 2 }) { "bad" }).toString(), "ok")
            __check((check(1, { it > 2 }) { "bad" }).toString(), "bad")
        }
