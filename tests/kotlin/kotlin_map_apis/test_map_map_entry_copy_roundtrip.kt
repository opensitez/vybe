// vybe-test: kotlin/kotlin_map_apis/test_map_map_entry_copy_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("x" to 9)
            val round = linkedMapOf(map.entries.first().toPair())
            __check((round["x"]).toString(), "9")
            __check((round.size).toString(), "1")
        }
