// vybe-test: kotlin/generics/test_generic_function_with_three_comparable_values
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T : Comparable<T>> maxOfThree(a: T, b: T, c: T): T {
            return if (a > b && a > c) a else if (b > c) b else c
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((maxOfThree(4, 9, 1)).toString(), "9")
            __check((maxOfThree("alpha", "gamma", "beta")).toString(), "gamma")
        }
