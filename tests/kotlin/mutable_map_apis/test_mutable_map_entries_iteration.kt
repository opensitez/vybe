// vybe-test: kotlin/mutable_map_apis/test_mutable_map_entries_iteration
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("x" to 1, "y" to 2)
            val out = values.entries.joinToString("|") { it.key + ":" + it.value }
            __check((out).toString(), "x:1|y:2")
        }
