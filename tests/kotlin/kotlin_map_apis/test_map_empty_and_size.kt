// vybe-test: kotlin/kotlin_map_apis/test_map_empty_and_size
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = emptyMap<String, Int>()
            __check((map.isEmpty()).toString(), "true")
            __check((map.isNotEmpty()).toString(), "false")
            __check((map.size).toString(), "0")
        }
