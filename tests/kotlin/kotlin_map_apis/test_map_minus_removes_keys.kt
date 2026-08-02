// vybe-test: kotlin/kotlin_map_apis/test_map_minus_removes_keys
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            val reduced = map - "b"
            __check((reduced.size).toString(), "2")
            __check((reduced.containsKey("b")).toString(), "false")
        }
