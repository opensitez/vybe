// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_duplicate_keys_in_to_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(Pair("k", 1), Pair("k", 2), Pair("k", 3))
            val map = values.toMap()
            __check((map.size).toString(), "1")
            __check((map["k"]).toString(), "3")
        }
