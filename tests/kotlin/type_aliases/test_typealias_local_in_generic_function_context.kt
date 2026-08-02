// vybe-test: kotlin/type_aliases/test_typealias_local_in_generic_function_context
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Wrapped<T> = List<T>

        fun <T> describe(values: Wrapped<T>): String {
            typealias FirstLabel = String
            val first: FirstLabel = values.firstOrNull().toString()
            return first
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val names = listOf("kotlin", "tests")
            __check((describe(names)).toString(), "kotlin")
            val numbers = listOf(1, 2, 3)
            __check((describe(numbers)).toString(), "1")
        }
