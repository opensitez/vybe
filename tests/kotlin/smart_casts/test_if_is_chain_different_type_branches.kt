// vybe-test: kotlin/smart_casts/test_if_is_chain_different_type_branches
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun classify(value: Any): String {
            return if (value is Int) {
                "int-" + value
            } else if (value is String) {
                "string-" + value.length
            } else {
                "other"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(9)).toString(), "int-9")
            __check((classify("abc")).toString(), "string-3")
            __check((classify(true)).toString(), "other")
        }
