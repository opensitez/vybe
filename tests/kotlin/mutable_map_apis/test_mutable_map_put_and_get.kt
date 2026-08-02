// vybe-test: kotlin/mutable_map_apis/test_mutable_map_put_and_get
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1)
            val prev = values.put("b", 2)
            __check((prev).toString(), "null")
            __check((values["b"]).toString(), "2")
        }
