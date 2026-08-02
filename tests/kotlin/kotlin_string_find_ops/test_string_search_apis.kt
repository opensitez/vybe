// vybe-test: kotlin/kotlin_string_find_ops/test_string_search_apis
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "banana"
            __check((s.indexOf("na")).toString(), "2")
            __check((s.lastIndexOf("na")).toString(), "4")
            __check((s.contains("an")).toString(), "true")
        }
