// vybe-test: kotlin/recursion/test_recursion_nested_while_simulation
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun dropCount(values: List<Int>): Int {
            return if (values.isEmpty()) 0 else 1 + dropCount(values.drop(1).drop(1))
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((dropCount(listOf(1, 2, 3, 4))).toString(), "2")
        }
