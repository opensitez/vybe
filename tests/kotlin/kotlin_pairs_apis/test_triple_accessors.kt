// vybe-test: kotlin/kotlin_pairs_apis/test_triple_accessors
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Triple("alpha", 11, true)
            __check((t.first).toString(), "alpha")
            __check((t.second).toString(), "11")
            __check((t.third).toString(), "true")
        }
