// vybe-test: kotlin/kotlin_type_parameter_bounds/test_generic_typealias_with_upper_bound
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

typealias ComparableNumber<T> = T where T : Comparable<T>, T : Number

        fun pick(a: ComparableNumber<Int>, b: ComparableNumber<Int>): Int = if (a > b) a else b

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pick(10, 4)).toString(), "10")
        }
