// vybe-test: kotlin/type_aliases/test_typealias_for_generic_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias BoxOfInt = MutableList<Int>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: BoxOfInt = mutableListOf(1, 2, 3)
            values.add(4)
            __check((values.joinToString(",")).toString(), "1,2,3,4")
        }
