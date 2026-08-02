// vybe-test: kotlin/kotlin_map_apis/test_map_replace_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("x" to 1)
            val previous = map.put("x", 7)
            __check((previous).toString(), "1")
            __check((map["x"]).toString(), "7")
        }
