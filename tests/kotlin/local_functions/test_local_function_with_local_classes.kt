// vybe-test: kotlin/local_functions/test_local_function_with_local_classes
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun make(): String {
                class Holder(val v: Int)
                return Holder(9).v.toString()
            }
            __check((make()).toString(), "9")
        }
