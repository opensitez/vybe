// vybe-test: kotlin/generics/test_generic_constraint_comparable_max
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T : Comparable<T>> maxOf(first: T, second: T): T {
            return if (first > second) first else second
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((maxOf(4, 9)).toString(), "9")
            __check((maxOf("a", "z")).toString(), "z")
        }
