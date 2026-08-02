// vybe-test: kotlin/type_aliases/test_typealias_for_mutable_set_specializes_numeric_key
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias NumberSet = MutableSet<Int>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: NumberSet = hashSetOf(1, 2, 2, 3)
            values.add(3)
            __check((values.size).toString(), "3")
            __check((values.contains(2)).toString(), "true")
        }
