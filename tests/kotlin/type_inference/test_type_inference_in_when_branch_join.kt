// vybe-test: kotlin/type_inference/test_type_inference_in_when_branch_join
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = when (1) {
                1 -> "one"
                else -> "other"
            }
            __check((value).toString(), "one")
        }
