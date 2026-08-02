// vybe-test: kotlin/type_aliases/test_typealias_for_generic_list_alias
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias StringList = List<String>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val names: StringList = listOf("a", "b", "c")
            __check((names.size).toString(), "3")
            __check((names[1]).toString(), "b")
        }
