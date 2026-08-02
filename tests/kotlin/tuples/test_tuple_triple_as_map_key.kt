// vybe-test: kotlin/tuples/test_tuple_triple_as_map_key
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf(Triple("a", 1, true) to "ok")
            __check((map[Triple("a", 1, true)]).toString(), "ok")
            __check((map[Triple("a", 1, false)] ?: "missing").toString(), "missing")
        }
