// vybe-test: kotlin/type_aliases/test_typealias_map_entry_is_a_type_projection_only
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias ScoreEntry = Map.Entry<String, Int>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 10, "b" to 20)
            val top: ScoreEntry = map.entries.reduce { acc, item ->
                if (item.value > acc.value) item else acc
            }
            __check((top.key).toString(), "b")
            __check((top.value).toString(), "20")
        }
