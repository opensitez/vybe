// vybe-test: kotlin/kotlin_string_find_contains/test_last_index_of
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_contains.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "banana"
            __check((s.lastIndexOf("na").toString()).toString(), "4")
            __check((s.lastIndexOf("x").toString()).toString(), "-1")
        }
