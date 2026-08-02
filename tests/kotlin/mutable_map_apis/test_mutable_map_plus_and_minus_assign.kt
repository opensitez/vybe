// vybe-test: kotlin/mutable_map_apis/test_mutable_map_plus_and_minus_assign
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1)
            values += mapOf("b" to 2)
            values -= "a"
            __check((values.keys.joinToString(",")).toString(), "b")
        }
