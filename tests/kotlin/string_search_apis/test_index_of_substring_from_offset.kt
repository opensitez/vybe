// vybe-test: kotlin/string_search_apis/test_index_of_substring_from_offset
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "banana"
            __check((text.indexOf("na")).toString(), "2")
            __check((text.indexOf("na", 3)).toString(), "4")
            __check((text.lastIndexOf("na")).toString(), "4")
        }
