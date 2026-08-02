// vybe-test: kotlin/mutable_set_apis/test_mutable_set_filter_by_string_prefix
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf("aa", "ab", "bb")
            val filtered = values.filter { it.startsWith("a") }.toMutableSet()
            __check((filtered.joinToString(",")).toString(), "aa,ab")
        }
