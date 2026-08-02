// vybe-test: kotlin/type_inference/test_type_inference_of_caller_callee_with_any
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun asAny(v: Any) = v.toString()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = asAny(9)
            __check((x).toString(), "9")
        }
