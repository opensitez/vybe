// vybe-test: kotlin/kotlin_string_find_contains/test_index_of_first
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_contains.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "banana"
            __check((s.indexOf("na").toString()).toString(), "2")
            __check((s.indexOf("na", startIndex = 3).toString()).toString(), "4")
        }
