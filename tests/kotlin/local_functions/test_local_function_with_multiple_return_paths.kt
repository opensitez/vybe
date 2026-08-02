// vybe-test: kotlin/local_functions/test_local_function_with_multiple_return_paths
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun classify(v: Int): String {
                if (v < 0) return "neg"
                if (v == 0) return "zero"
                return "pos"
            }
            __check((classify(-1)).toString(), "neg")
            __check((classify(0)).toString(), "zero")
            __check((classify(1)).toString(), "pos")
        }
