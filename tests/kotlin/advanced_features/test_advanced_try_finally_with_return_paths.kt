// vybe-test: kotlin/advanced_features/test_advanced_try_finally_with_return_paths
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

var marker = ""

        fun evaluate(use_fast: Boolean): Int {
            try {
                if (use_fast) {
                    return 7
                } else {
                    return 11
                }
            } finally {
                marker += "f"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((evaluate(true)).toString(), "7")
            __check((evaluate(false)).toString(), "11")
            __check((marker).toString(), "ff")
        }
