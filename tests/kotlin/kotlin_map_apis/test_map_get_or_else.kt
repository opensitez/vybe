// vybe-test: kotlin/kotlin_map_apis/test_map_get_or_else
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1)
            __check((map.getOrElse("a") { 0 }).toString(), "1")
            __check((map.getOrElse("z") { 7 }).toString(), "7")
        }
