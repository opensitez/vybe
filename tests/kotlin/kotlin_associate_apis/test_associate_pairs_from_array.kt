// vybe-test: kotlin/kotlin_associate_apis/test_associate_pairs_from_array
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val entries = arrayOf(Pair("a", 1), Pair("b", 2), Pair("c", 3))
            val map = entries.toMap()
            __check((map.keys.joinToString(",")).toString(), "a,b,c")
            __check((map.values.sum()).toString(), "6")
        }
