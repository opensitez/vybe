// vybe-test: kotlin/string_search_apis/test_replace_and_replace_first_boundaries
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "banana"
            __check((text.replaceFirst("ba", "pa")).toString(), "panana")
            __check((text.replace("na", "xx")).toString(), "baxxxx")
            __check((text.replace("na", "xx", false)).toString(), "baxxxx")
        }
