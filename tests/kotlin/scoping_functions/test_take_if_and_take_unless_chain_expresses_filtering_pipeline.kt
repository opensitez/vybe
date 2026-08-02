// vybe-test: kotlin/scoping_functions/test_take_if_and_take_unless_chain_expresses_filtering_pipeline
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Box(var n: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Box(7)
            val result = value
                .takeIf { it.n > 5 }
                ?.takeUnless { it.n % 2 == 0 }
                ?.n ?: -1
            __check((result).toString(), "7")
        }
