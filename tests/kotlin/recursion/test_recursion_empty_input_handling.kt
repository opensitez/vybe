// vybe-test: kotlin/recursion/test_recursion_empty_input_handling
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun total(values: List<Int>): Int {
            if (values.isEmpty()) return 0
            return values.first() + total(values.drop(1))
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((total(listOf<Int>())).toString(), "0")
            __check((total(listOf(9))).toString(), "9")
        }
