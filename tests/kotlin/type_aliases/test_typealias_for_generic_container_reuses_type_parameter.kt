// vybe-test: kotlin/type_aliases/test_typealias_for_generic_container_reuses_type_parameter
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Container<T> = MutableList<T>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: Container<Int> = mutableListOf(1, 2)
            values.add(3)
            __check((values.joinToString(",")).toString(), "1,2,3")
        }
