// vybe-test: kotlin/mutable_map_apis/test_mutable_map_merge_like_update
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1)
            val next = values.getOrElse("a") { 0 } + 4
            values["a"] = next
            __check((values["a"]).toString(), "5")
        }
