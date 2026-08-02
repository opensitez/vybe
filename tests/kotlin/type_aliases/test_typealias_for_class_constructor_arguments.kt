// vybe-test: kotlin/type_aliases/test_typealias_for_class_constructor_arguments
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

data class PairValue(val id: Int, val label: String)
        typealias Entry = PairValue

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Entry = Entry(3, "x")
            __check((item.id).toString(), "3")
            __check((item.label).toString(), "x")
        }
