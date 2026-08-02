// vybe-test: kotlin/kotlin_map_apis/test_map_with_pair_array_to_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pairs = arrayOf(Pair("n", 1), Pair("m", 2), Pair("n", 4))
            val map = pairs.toMap()
            __check((map.size).toString(), "2")
            __check((map["n"]).toString(), "4")
        }
