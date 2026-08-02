// vybe-test: kotlin/kotlin_map_apis/test_map_remove_key
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            val removed = map.remove("b")
            __check((removed).toString(), "2")
            __check((map.size).toString(), "2")
            __check((map.containsKey("b")).toString(), "false")
        }
