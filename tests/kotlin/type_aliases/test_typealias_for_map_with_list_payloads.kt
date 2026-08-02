// vybe-test: kotlin/type_aliases/test_typealias_for_map_with_list_payloads
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias ScoresByLabel = MutableMap<String, MutableList<Int>>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val scores: ScoresByLabel = mutableMapOf()
            scores["a"] = mutableListOf(1, 2)
            scores["a"]?.add(3)
            __check((scores["a"]?.size).toString(), "3")
            __check((scores["a"]?.sum()).toString(), "6")
        }
