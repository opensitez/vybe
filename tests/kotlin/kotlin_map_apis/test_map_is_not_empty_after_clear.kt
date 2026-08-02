// vybe-test: kotlin/kotlin_map_apis/test_map_is_not_empty_after_clear
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1)
            val before = map.isNotEmpty()
            map.clear()
            __check((before).toString(), "true")
            __check((map.isEmpty()).toString(), "true")
        }
