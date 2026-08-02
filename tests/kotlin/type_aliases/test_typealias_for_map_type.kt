// vybe-test: kotlin/type_aliases/test_typealias_for_map_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias StringNumberMap = Map<String, Int>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: StringNumberMap = mapOf("a" to 1, "b" to 2)
            __check((values["a"]).toString(), "1")
            __check((values["c"] == null).toString(), "true")
        }
