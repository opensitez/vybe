// vybe-test: kotlin/generics/test_generic_covariant_readonly_collection_can_receive_concrete_subtype
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> countValues(values: List<T>): Int {
            return values.size
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ints: MutableList<Int> = mutableListOf(1, 2, 3)
            val numbers: List<Number> = ints
            __check((countValues(numbers)).toString(), "3")
            __check((countValues(ints)).toString(), "3")
        }
