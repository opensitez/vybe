// vybe-test: kotlin/mutable_map_apis/test_mutable_map_keys_iteration
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("x" to 1, "y" to 2)
            val joined = values.keys.joinToString(",")
            __check((joined).toString(), "x,y")
        }
