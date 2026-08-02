// vybe-test: kotlin/type_aliases/test_typealias_for_mutable_map_key_and_value_types
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias StringToIntMap = MutableMap<String, Int>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: StringToIntMap = mutableMapOf("a" to 1)
            values["b"] = 2
            __check((values["a"]).toString(), "1")
            __check((values["b"]).toString(), "2")
        }
