// vybe-test: kotlin/generics/test_generic_function_with_multiple_return_types
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <A, B> pairLabel(left: A, right: B): String {
            return left.toString() + ":" + right.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pairLabel(true, 1)).toString(), "true:1")
            __check((pairLabel(2.2, "x")).toString(), "2.2:x")
            __check((pairLabel("k", false)).toString(), "k:false")
        }
