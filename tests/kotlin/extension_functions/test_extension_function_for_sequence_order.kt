// vybe-test: kotlin/extension_functions/test_extension_function_for_sequence_order
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun Int.next(): Int = this + 1
        fun Int.prev(): Int = this - 1

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = 4
            __check((base.next().prev().next()).toString(), "5")
        }
