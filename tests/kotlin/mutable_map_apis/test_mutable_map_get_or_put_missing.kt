// vybe-test: kotlin/mutable_map_apis/test_mutable_map_get_or_put_missing
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1)
            val value = values.getOrPut("b") { 8 }
            __check((value).toString(), "8")
            __check((values["b"]).toString(), "8")
        }
